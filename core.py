"""Lógica compartida de Facturación FIN.ARQ (versión 2).

Copia sin cambios de las reglas de LayOut.py: constantes, lectura de XML CFDI y
estados de cuenta, y la capa de datos (Google Sheets con respaldo en CSV local).
La interfaz vive en app_v2.py; aquí no se dibuja nada en pantalla.
"""
import io
import os
import calendar
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
import unicodedata
import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
import re
import pdfplumber


IVA_RATE = 0.16

ASESORES = ["Jorge Orozco", "Andrea Dávila", "Ximena Mora",
            "Marco Ochoa", "Eduardo Jáquez", "Negocios-FINARQ"]

PRODUCTOS = ["FX", "Terminal Punto de Venta", "Crédito",
             "Arrendamiento", "Honorarios / Asesoría", "Credito TPV",
             "Inversión", "Factoraje"]

ESTATUS = ["LISTO", "TIMBRADA", "PAGADA", "CANCELADA"]

# VENCIDA solo se muestra en pantalla (pago programado vencido y no PAGADA ni CANCELADA); no se guarda
ESTATUS_VISTA = ["VENCIDA"] + ESTATUS
ESTATUS_ORDEN = {e: i for i, e in enumerate(ESTATUS_VISTA)}

MESES_CORTO = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
               "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]

COLUMNS = [
    "CÓDIGO SAT", "CLIENTE", "CONCEPTO", "PRODUCTO", "Aliado Comercial", "Asesor",
    "BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL",
    "FOLIO (LISTO P/FACTURA)", "FECHA DE FACTURACION", "ESTATUS DE PAGO",
    "FECHA DE PAGO", "NOTAS",
    "Datos (Factura)", "Concepto (Factura)", "Número Factura",
    "Fecha Emisión", "Método de Pago",
    "Complemento de Pago", "Pago Real", "Comprobante de Pago", "Comentarios",
    "ID", "Capturado_Por", 
]

# Columnas que llegan de la factura 
FACTURA_COLS = ["Datos (Factura)", "Concepto (Factura)", "Número Factura",
                "Fecha Emisión", "Método de Pago"]

ESTADO_COLS = ["Complemento de Pago", "Pago Real", "Comprobante de Pago", "Comentarios"]

LAYOUT_COLS = [c for c in COLUMNS if c not in FACTURA_COLS + ESTADO_COLS]

COMPLETO_COLS = ["CÓDIGO SAT", "CLIENTE", "CONCEPTO", "PRODUCTO", "Aliado Comercial",
                 "Asesor", "BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL"]

EDITABLE_COLS = [c for c in LAYOUT_COLS if c not in ("ID", "Capturado_Por")]

MONEY_COLS = ["BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL"]

UPPER_COLS = ["CÓDIGO SAT", "CLIENTE", "CONCEPTO", "PRODUCTO", "Aliado Comercial", "Asesor",
              "ESTATUS DE PAGO", "NOTAS", "Datos (Factura)", "Concepto (Factura)",
              "Número Factura", "Método de Pago", "Complemento de Pago",
              "Comprobante de Pago", "Comentarios"]

ESTADO_CUENTA_COLS = ["Complemento de Pago", "Pago Real", "Comprobante de Pago", "Comentarios"]

MASTER_COLS = [
    "Mes", "Cliente/Proveedor", "Datos", "Asesor", "Producto", "Concepto",
    "Volumen", "Subtotal", "IVA", "Total",
    "Clave SAT", "Factura", "Emisión", "Método de Pago", "Pago Programado", "Fecha de Pago",
    "Complemento de Pago", "Comprobante de Pago", "Comentarios",
]

MASTER_FACT_COLS = [c for c in MASTER_COLS if c not in ESTADO_CUENTA_COLS]


# Valores predeterminados para FX
FX_ALIADO = "INVEX"
FX_CODIGO_SAT = "84121701"

# Valores predeterminados para registros creados desde XML (Kushki)
KUSHKI_ALIADO = "KUSHKI"
KUSHKI_PRODUCTO = "Terminal Punto de Venta"
KUSHKI_ASESOR = "Negocios-FINARQ"
CLIENTES_FX_DEFAULT = [
    "FUNNY KITCHEN", "GRUPO CEBANDI", "VALA PADEL",
    "AGRICULTURA Y GANADERIA DEL BAJIO DE RATONES", "Base Ilustre México",
    "Tequila Cascahuín", "DESTILERIA EL PANDILLO", "COMERCIALIZADORA HEXCENTRIX",
    "JORGE ENRIQUE CARRILLO SAMWAYS", "PRIMO SPIRITS", "AGROPLUMA",
    "MAQUILADORA DE SERVICIOS MSM", "Gabriela Chavez Ashida", "LUIS LLAMA",
    "COMERCIALIZADORA IBK", "ABUNDANCIA EN ACCION", "CARLOS GONZALEZ JAIDAR",
    "ERAUNBOY", "Culture Tea", "Luna Tea", "BIOTECNOLOGIA EN CRECIMIENTO",
    "ROGELIO GUTIERREZ FRAGA",
]



#LEER CFDI XML

def _norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return "".join(ch for ch in s.upper() if ch.isalnum())


def _tokens(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return [t for t in "".join(ch if ch.isalnum() else " " for ch in s).split() if len(t) > 2]


def match_concepto(cliente, conceptos) -> str:
    conceptos = [c for c in conceptos if c]
    if not conceptos:
        return ""
    if len(conceptos) == 1:
        return conceptos[0]
    nc = _norm(cliente)
    if nc:
        for d in conceptos:
            if nc in _norm(d):
                return d
        ct = set(_tokens(cliente))
        best, score = "", 0
        for d in conceptos:
            s = len(ct & set(_tokens(d)))
            if s > score:
                best, score = d, s
        if score > 0:
            return best
    return " | ".join(conceptos)


def parse_cfdi(xml_bytes: bytes) -> dict:
    """Extrae de un XML CFDI (3.3/4.0) los campos de la factura."""
    c = ET.fromstring(xml_bytes)
    cfdi_ns = c.tag.split("}")[0].strip("{")
    ns = {"cfdi": cfdi_ns, "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital"}
    rec = c.find("cfdi:Receptor", ns)
    rfc = rec.get("Rfc", "") if rec is not None else ""
    nombre = rec.get("Nombre", "") if rec is not None else ""
    nodos = c.findall("cfdi:Conceptos/cfdi:Concepto", ns)
    conceptos = [cc.get("Descripcion", "") for cc in nodos]
    serie = c.get("Serie") or ""
    folio = c.get("Folio") or ""
    num = (serie + "-" if serie else "") + folio
    tfd = c.find(".//tfd:TimbreFiscalDigital", ns)
    imp = c.find("cfdi:Impuestos", ns)
    return {
        "Datos (Factura)": ", ".join([x for x in (nombre, rfc) if x]),
        "Concepto (Factura)": " | ".join([x for x in conceptos if x]),
        "Número Factura": num,
        "Fecha Emisión": (c.get("Fecha") or "")[:10],
        "Método de Pago": c.get("MetodoPago") or "",
        "_Total": c.get("Total") or "",
        "_Subtotal": c.get("SubTotal") or c.get("Subtotal") or "",
        "_Descuento": c.get("Descuento") or "",
        "_IVA": (imp.get("TotalImpuestosTrasladados") if imp is not None else "") or "",
        "_UUID": tfd.get("UUID") if tfd is not None else "",
        "_Conceptos": [x for x in conceptos if x],
        "_Nombre": nombre,
        "_ClaveSAT": nodos[0].get("ClaveProdServ", "") if nodos else "",
        "_Tipo": (c.get("TipoDeComprobante") or "").upper(),
    }

def calcular_pago_programado(metodo, fecha_emision) -> str:
    try:
        d = datetime.strptime(str(fecha_emision)[:10], "%Y-%m-%d").date()
    except Exception:
        return ""
    m = str(metodo).strip().upper()
    if m == "PUE":
        return d.replace(day=calendar.monthrange(d.year, d.month)[1]).isoformat()
    if m == "PPD":
        return (d + timedelta(days=60)).isoformat()
    return ""


def layout_a_master(df: pd.DataFrame) -> pd.DataFrame:
    #Arma la tabla del Registro Máster con lo que hoy se obtiene del LayOut.
    out = pd.DataFrame(columns=MASTER_COLS)
    if df.empty:
        return out
    _f = pd.to_datetime(df["FECHA DE FACTURACION"], errors="coerce")
    out["Mes"] = _f.map(lambda d: f"{MESES_CORTO[d.month - 1]}-{d.year}" if pd.notna(d) else "")
    out["Cliente/Proveedor"] = df["CLIENTE"].values
    out["Asesor"] = df["Asesor"].values
    out["Producto"] = df["PRODUCTO"].values
    out["Volumen"] = df["BASE (Volumen)"].values
    out["Subtotal"] = df["SUBTOTAL (INGRESO)"].values
    out["IVA"] = df["IVA"].values
    out["Total"] = df["TOTAL"].values
    out["Clave SAT"] = df["CÓDIGO SAT"].values
    # Datos que llegan de la FACTURA (XML CFDI)
    out["Datos"] = df["Datos (Factura)"].values
    out["Concepto"] = df["Concepto (Factura)"].values
    out["Factura"] = df["Número Factura"].values
    out["Emisión"] = df["Fecha Emisión"].values
    out["Método de Pago"] = df["Método de Pago"].values
    out["Pago Programado"] = [calcular_pago_programado(m, f) for m, f in zip(df["Método de Pago"].values, df["Fecha Emisión"].values)]
    out["Fecha de Pago"] = df["FECHA DE PAGO"].values
    out["Complemento de Pago"] = df["Complemento de Pago"].values
    out["Comprobante de Pago"] = df["Comprobante de Pago"].values
    out["Comentarios"] = df["Comentarios"].values
    return out.fillna("")

MESES_ABR = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
             "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}
_AMT_RE = re.compile(r"HORA\s+\d{1,2}:\d{2}\s+SUC\s+\d+\s+([\d,]+\.\d{2})(?:\s+([\d,]+\.\d{2}))?")
_DATE_RE = re.compile(r"^(\d{1,2})\s+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(.*)")


def parse_estado_cuenta(pdf_bytes: bytes) -> list:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        txt = "\n".join((p.extract_text() or "") for p in pdf.pages)
    my = re.search(r"ESTADO DE CUENTA AL\s+\d{1,2}\s+DE\s+\w+\s+DE\s+(\d{4})", txt)
    anio = my.group(1) if my else str(date.today().year)
    trans, cur = [], None
    for line in txt.split("\n"):
        s = line.strip()
        md = _DATE_RE.match(s)
        ma = _AMT_RE.search(line)
        if md:
            if cur:
                trans.append(cur)
            dd, mm = int(md.group(1)), MESES_ABR[md.group(2)]
            cur = {"fecha": f"{anio}-{mm:02d}-{dd:02d}", "concepto": [md.group(3)], "monto": None}
        elif ma and cur and cur["monto"] is None:
            cur["monto"] = float(ma.group(1).replace(",", ""))
        elif cur and cur["monto"] is None:
            if (s and not s.startswith("CAJA") and "DETALLE DE OPER" not in s
                    and "ESTADO DE CUENTA" not in s and not s.startswith("Página")
                    and not s.startswith("FECHA CONCEPTO")):
                cur["concepto"].append(s)
    if cur:
        trans.append(cur)
    deps = []
    for d in trans:
        tc = " ".join(d["concepto"])
        if "PAGO RECIBIDO" in tc.upper() and d["monto"]:
            deps.append({"fecha": d["fecha"], "monto": d["monto"], "concepto": tc})
    return deps


def match_deposito(total, deps):
    try:
        t = round(float(total), 2)
    except (ValueError, TypeError):
        return None
    for d in deps:
        if abs(d["monto"] - t) < 0.01:
            return d
    return None

def _fecha_iso(v) -> str:
    f = pd.to_datetime(str(v).strip() or None, errors="coerce")
    return f.date().isoformat() if pd.notna(f) else ""


def montos_xml(datos):
    total = pd.to_numeric(datos.get("_Total"), errors="coerce")
    if pd.isna(total):
        return None
    sub = pd.to_numeric(datos.get("_Subtotal"), errors="coerce")
    desc = pd.to_numeric(datos.get("_Descuento"), errors="coerce")
    iva = pd.to_numeric(datos.get("_IVA"), errors="coerce")
    if pd.notna(sub) and pd.notna(desc):
        sub = sub - desc
    if pd.isna(sub):
        sub = round(float(total) / (1 + IVA_RATE), 2)
    if pd.isna(iva):
        iva = round(float(total) - float(sub), 2)
    return round(float(sub), 2), round(float(iva), 2), round(float(total), 2)



def aplicar_factura(reg: pd.Series, datos: dict):
    """Devuelve (cambios, corregido): las columnas a escribir en el registro al ligar la factura."""
    cambios = {col: datos[col] for col in FACTURA_COLS}
    cambios["FECHA DE FACTURACION"] = datos["Fecha Emisión"]
    cambios["FOLIO (LISTO P/FACTURA)"] = datos["Número Factura"]

    corregido = False
    m = montos_xml(datos)
    reg_total = pd.to_numeric(reg["TOTAL"], errors="coerce")
    if m is not None and (pd.isna(reg_total) or abs(m[2] - float(reg_total)) > 0.01):
        cambios.update({"SUBTOTAL (INGRESO)": m[0], "IVA": m[1], "TOTAL": m[2]})
        corregido = True
    return cambios, corregido

def extraer_xmls(archivos) -> list:
    """Devuelve [(nombre, bytes)] de los XML subidos sueltos o dentro de ZIPs."""
    out = []
    for f in archivos:
        if f.name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(f.getvalue())) as z:
                for n in z.namelist():
                    base = os.path.basename(n)
                    if (n.lower().endswith(".xml") and not n.startswith("__MACOSX")
                            and not base.startswith("._")):
                        out.append((base, z.read(n)))
        else:
            out.append((f.name, f.getvalue()))
    return out


def fila_desde_xml(datos: dict, rid: str, capturo: str) -> dict:
    """Arma un registro nuevo (Kushki) con lo que trae el XML."""
    m = montos_xml(datos) or ("", "", "")
    fila = {c: "" for c in COLUMNS}
    fila.update({c: datos[c] for c in FACTURA_COLS})
    fila.update({
        "CÓDIGO SAT": datos["_ClaveSAT"], "CLIENTE": datos["_Nombre"],
        "CONCEPTO": datos["Concepto (Factura)"],
        "PRODUCTO": KUSHKI_PRODUCTO, "Aliado Comercial": KUSHKI_ALIADO,
        "Asesor": KUSHKI_ASESOR,
        "SUBTOTAL (INGRESO)": m[0], "IVA": m[1], "TOTAL": m[2],
        "FOLIO (LISTO P/FACTURA)": datos["Número Factura"],
        "FECHA DE FACTURACION": datos["Fecha Emisión"],
        "ESTATUS DE PAGO": "TIMBRADA",
        "ID": rid, "Capturado_Por": f"{capturo} (XML)",
    })
    return fila



# CAPA DE DATOS  (Google Sheets con fallback a CSV local)
CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "layout.csv")

CLIENTES_FX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "clientes_fx.txt")


def _clientes_fx_locales() -> list:
    if os.path.exists(CLIENTES_FX_PATH):
        with open(CLIENTES_FX_PATH, encoding="utf-8") as f:
            xs = [l.strip() for l in f if l.strip()]
        return xs or list(CLIENTES_FX_DEFAULT)
    return list(CLIENTES_FX_DEFAULT)


@st.cache_resource(show_spinner=False)
def _get_ws_clientes():
    """Pestaña «ClientesFX» del mismo Google Sheet (se crea con la lista inicial si no existe)."""
    import gspread
    sh = _get_worksheet().spreadsheet
    try:
        ws = sh.worksheet("ClientesFX")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="ClientesFX", rows=500, cols=1)
    if len(ws.col_values(1)) <= 1:
        ws.update([["CLIENTE"]] + [[c] for c in _clientes_fx_locales()], "A1")
    return ws


@st.cache_data(ttl=300, show_spinner=False)
def cargar_clientes_fx() -> list:
    if _usar_sheets():
        xs = [v.strip() for v in _get_ws_clientes().col_values(1)[1:] if v.strip()]
        return xs or list(CLIENTES_FX_DEFAULT)
    return _clientes_fx_locales()


def agregar_cliente_fx(nombre: str):
    nombre = nombre.strip()
    xs = cargar_clientes_fx()
    if nombre and nombre not in xs:
        if _usar_sheets():
            _get_ws_clientes().append_row([nombre])
        else:
            os.makedirs(os.path.dirname(CLIENTES_FX_PATH), exist_ok=True)
            with open(CLIENTES_FX_PATH, "w", encoding="utf-8") as f:
                f.write("\n".join(xs + [nombre]))
        cargar_clientes_fx.clear()


def _usar_sheets() -> bool:
    try:
        return "gcp_service_account" in st.secrets and "sheet_url" in st.secrets
    except Exception:
        return False


def url_sheet() -> str:
    """Liga a la hoja de Google Sheets (para el botón de la barra lateral)."""
    url = str(st.secrets.get("sheet_url", "")).strip()
    if url.startswith("http"):
        return url
    return f"https://docs.google.com/spreadsheets/d/{st.secrets.get('sheet_id', url)}"


@st.cache_resource(show_spinner=False)
def _get_worksheet():
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets"
    ]

    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=scopes
        )
        gc = gspread.authorize(creds)

    except Exception as e:
        st.error("ERROR EN AUTENTICACIÓN")
        st.exception(e)
        raise

    try:
        sh = gc.open_by_key(st.secrets["sheet_id"])

    except PermissionError as e:
        st.error("ERROR AL ABRIR GOOGLE SHEETS")
        st.write("Cuenta utilizada:", creds.service_account_email)
        st.write("Proyecto:", creds.project_id)
        st.write("Sheet ID:", st.secrets["sheet_id"])

        # Intentar mostrar la causa original
        if e.__cause__:
            st.write("Error original:")
            st.exception(e.__cause__)

        raise

    try:
        ws = sh.worksheet("LayOut")

    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title="LayOut",
            rows=1000,
            cols=len(COLUMNS)
        )
        ws.append_row(COLUMNS)

    if ws.row_values(1) != COLUMNS:
        ws.update([COLUMNS], "A1")

    return ws


@st.cache_data(ttl=60, show_spinner="Leyendo registros…")
def leer_registros() -> pd.DataFrame:
    if _usar_sheets():
        ws = _get_worksheet()
        data = ws.get_all_records()
        df = pd.DataFrame(data, columns=COLUMNS) if data else pd.DataFrame(columns=COLUMNS)
    else:
        if os.path.exists(CSV_PATH):
            df = pd.read_csv(CSV_PATH, dtype=str)
        else:
            df = pd.DataFrame(columns=COLUMNS)
    df = df.reindex(columns=COLUMNS)
    for c in UPPER_COLS:                     
        df[c] = df[c].fillna("").astype(str).str.upper().replace("NAN", "")
    df["ESTATUS DE PAGO"] = df["ESTATUS DE PAGO"].replace({"PENDIENTE": "LISTO", "": "LISTO"})
    # FECHA DE PAGO y Pago Real son el mismo dato: se mantienen iguales
    _pr = df["Pago Real"].fillna("").astype(str).str.strip().replace({"nan": ""})
    _fp = df["FECHA DE PAGO"].fillna("").astype(str).str.strip().replace({"nan": ""})
    df["FECHA DE PAGO"] = _pr.where(_pr != "", _fp)
    df["Pago Real"] = df["FECHA DE PAGO"]
    return df
       


# Las escrituras tocan solo las filas/celdas necesarias (nunca se borra y reescribe la hoja
# completa), y ubican cada registro por su ID leyendo la hoja en ese momento, sin caché.

def _celda(v) -> str:
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return str(v)


def _igual(a, b) -> bool:
    va, vb = _celda(a).strip(), _celda(b).strip()
    if va == vb:
        return True
    try:
        return abs(float(va) - float(vb)) < 0.005
    except ValueError:
        return False


def _leer_csv() -> pd.DataFrame:
    if os.path.exists(CSV_PATH):
        return pd.read_csv(CSV_PATH, dtype=str).reindex(columns=COLUMNS).fillna("")
    return pd.DataFrame(columns=COLUMNS)


def _escribir_csv(df: pd.DataFrame):
    os.makedirs(os.path.dirname(CSV_PATH), exist_ok=True)
    df.to_csv(CSV_PATH, index=False)


def _filas_por_id(ws) -> dict:
    """ID -> número de fila en la hoja."""
    ids = ws.col_values(COLUMNS.index("ID") + 1)[1:]
    return {str(v).strip(): i + 2 for i, v in enumerate(ids) if str(v).strip()}


def agregar_registros(filas: list, value_input_option: str = "RAW"):
    """Agrega registros nuevos al final, sin tocar los existentes."""
    rows = [[_celda(f.get(c, "")) for c in COLUMNS] for f in filas]
    if not rows:
        return
    if _usar_sheets():
        _get_worksheet().append_rows(rows, value_input_option=value_input_option)
    else:
        _escribir_csv(pd.concat([_leer_csv(), pd.DataFrame(rows, columns=COLUMNS)], ignore_index=True))
    leer_registros.clear()


def guardar_registro(fila: dict):
    """Agrega un registro nuevo al backend (Sheets o CSV)."""
    agregar_registros([{c: str(fila.get(c, "")).upper() for c in COLUMNS}], "USER_ENTERED")


def actualizar_registros(cambios: dict) -> list:
    """Escribe solo las celdas que cambiaron. cambios = {ID: {columna: valor}}.
    Devuelve los IDs que ya no existen (p. ej. alguien más los borró)."""
    cambios = {rid: c for rid, c in cambios.items() if c}
    if not cambios:
        return []
    if _usar_sheets():
        from gspread.utils import rowcol_to_a1
        ws = _get_worksheet()
        filas = _filas_por_id(ws)
        datos = [{"range": rowcol_to_a1(filas[rid], COLUMNS.index(col) + 1), "values": [[_celda(v)]]}
                 for rid, cols in cambios.items() if rid in filas for col, v in cols.items()]
        if datos:
            ws.batch_update(datos)
        faltan = [rid for rid in cambios if rid not in filas]
    else:
        df = _leer_csv()
        faltan = [rid for rid in cambios if rid not in set(df["ID"])]
        for rid, cols in cambios.items():
            idx = df.index[df["ID"] == rid]
            for col, v in cols.items():
                df.loc[idx, col] = _celda(v)
        _escribir_csv(df)
    leer_registros.clear()
    return faltan


def eliminar_registros(ids: list):
    """Borra solo las filas de esos IDs."""
    if not ids:
        return
    if _usar_sheets():
        ws = _get_worksheet()
        filas = _filas_por_id(ws)
        # De abajo hacia arriba, en una sola llamada, para que los números de fila no se recorran
        reqs = [{"deleteDimension": {"range": {"sheetId": ws.id, "dimension": "ROWS",
                                               "startIndex": n - 1, "endIndex": n}}}
                for n in sorted((filas[i] for i in ids if i in filas), reverse=True)]
        if reqs:
            ws.spreadsheet.batch_update({"requests": reqs})
    else:
        df = _leer_csv()
        _escribir_csv(df[~df["ID"].isin(ids)])
    leer_registros.clear()


def siguiente_id(df: pd.DataFrame) -> str:
    if df.empty or df["ID"].dropna().empty:
        return "MOV-00001"
    nums = pd.to_numeric(df["ID"].str.replace("MOV-", "", regex=False), errors="coerce").dropna()
    n = int(nums.max()) + 1 if not nums.empty else 1
    return f"MOV-{n:05d}"

def fmt_money(v) -> str:
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return str(v)

def es_completo(fila) -> bool:
    for c in COMPLETO_COLS:
        v = fila.get(c, "")
        if v is None or str(v).strip() == "" or str(v).strip().lower() == "nan":
            return False
    return True
 
 
def col_estado(df: pd.DataFrame) -> pd.Series:
    return df.apply(lambda r: "🟢" if es_completo(r) else "🔴", axis=1)

def pago_programado(df: pd.DataFrame) -> pd.Series:
    return pd.Series([calcular_pago_programado(m, f)
                      for m, f in zip(df["Método de Pago"].values, df["Fecha Emisión"].values)],
                     index=df.index)


def estatus_vista(df: pd.DataFrame) -> pd.Series:
    """ESTATUS DE PAGO con VENCIDA cuando el pago programado ya venció y no está PAGADA ni CANCELADA."""
    est = df["ESTATUS DE PAGO"].astype(str).str.upper()
    pp = pago_programado(df)
    vencido = (pp != "") & (pp < date.today().isoformat()) & ~est.isin(["PAGADA", "CANCELADA"])
    return est.where(~vencido, "VENCIDA")


def agrupar_por_mes(d: pd.DataFrame) -> pd.DataFrame:
    """Ordena por mes de facturación (reciente primero, sin fecha hasta arriba) y por estatus,
    con una fila en blanco entre meses."""
    if d.empty:
        return d
    f = pd.to_datetime(d["FECHA DE FACTURACION"], errors="coerce")
    mes = (f.dt.year * 100 + f.dt.month).fillna(999999)
    orden = d["ESTATUS DE PAGO"].map(ESTATUS_ORDEN).fillna(len(ESTATUS_ORDEN))
    d = d.assign(_mes=mes, _ord=orden).sort_values(["_mes", "_ord"], ascending=[False, True],
                                                  kind="stable")
    blanco = {c: (False if pd.api.types.is_bool_dtype(d[c])
                  else float("nan") if pd.api.types.is_numeric_dtype(d[c]) else "")
              for c in d.columns}
    partes = []
    for i, (_, g) in enumerate(d.groupby("_mes", sort=False)):
        if i:
            partes.append(pd.DataFrame([blanco]))
        partes.append(g)
    return pd.concat(partes, ignore_index=True).drop(columns=["_mes", "_ord"])


def _color_fila(row):
    est = str(row.get("ESTATUS DE PAGO", "")).upper()
    if est == "VENCIDA":
        css = "background-color: rgba(239, 68, 68, 0.30)"
    elif est == "LISTO":
        css = "background-color: rgba(59, 130, 246, 0.30)"   
    elif est == "PAGADA":
        css = "background-color: rgba(34, 197, 94, 0.30)"    
    else:
        css = ""                                             
    return [css] * len(row)

