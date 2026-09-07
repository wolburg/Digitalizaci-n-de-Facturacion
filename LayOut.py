import io
import os
import calendar
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
import unicodedata
import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
import re
import pdfplumber


# ----------------------------------------------------------------------------
# CONFIG GENERAL
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Facturación FIN.ARQ — LayOut", page_icon="🧾", layout="wide")

IVA_RATE = 0.16

ASESORES = ["Jorge Orozco", "Andrea Dávila", "Ximena Mora",
            "Marco Ochoa", "Eduardo Jáquez", "Negocios-FINARQ"]

PRODUCTOS = ["FX", "Terminal Punto de Venta", "Crédito",
             "Arrendamiento", "Honorarios / Asesoría", "Credito TPV"]

ESTATUS = ["PENDIENTE", "TIMBRADA", "PAGADA"]

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

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
    "Clave SAT", "Factura", "Emisión", "Método de Pago", "Pago Programado",
    "Complemento de Pago", "Pago Real", "Comprobante de Pago", "Comentarios",
]

MASTER_FACT_COLS = [c for c in MASTER_COLS if c not in ESTADO_CUENTA_COLS]


# Valores predeterminados para FX
FX_ALIADO = "INVEX"
FX_CODIGO_SAT = "84121701"
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
    conceptos = [cc.get("Descripcion", "") for cc in c.findall("cfdi:Conceptos/cfdi:Concepto", ns)]
    serie = c.get("Serie") or ""
    folio = c.get("Folio") or ""
    num = (serie + "-" if serie else "") + folio
    tfd = c.find(".//tfd:TimbreFiscalDigital", ns)
    return {
        "Datos (Factura)": ", ".join([x for x in (nombre, rfc) if x]),
        "Concepto (Factura)": " | ".join([x for x in conceptos if x]),
        "Número Factura": num,
        "Fecha Emisión": (c.get("Fecha") or "")[:10],
        "Método de Pago": c.get("MetodoPago") or "",
        "_Total": c.get("Total") or "",
        "_UUID": tfd.get("UUID") if tfd is not None else "",
        "_Conceptos": [x for x in conceptos if x],
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
    out["Mes"] = _f.dt.month.map(lambda n: MESES[int(n) - 1] if pd.notna(n) else "")
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
    out["Complemento de Pago"] = df["Complemento de Pago"].values
    out["Pago Real"] = df["Pago Real"].values
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

def render_facturacion(df: pd.DataFrame):
    st.header("📄 Facturación")
    st.markdown("### 📎 Adjuntar factura (XML CFDI) a un registro")
    if df.empty:
        st.info("Primero captura movimientos en el LayOut.")
    else:
        etiquetas = {r["ID"]: f'{r["ID"]} — {r["CLIENTE"]} · {fmt_money(r["TOTAL"])}'
                     for _, r in df.iterrows()}
        ca, cb = st.columns([1.2, 1])
        sel_id = ca.selectbox("Registro", options=list(etiquetas.keys()),
                              format_func=lambda x: etiquetas[x])
        xml_file = cb.file_uploader("XML de la factura", type=["xml"])
        if xml_file is not None:
            try:
                datos = parse_cfdi(xml_file.getvalue())
            except Exception as ex:
                st.error(f"No se pudo leer el XML: {ex}")
            else:
                cliente_reg = str(df.loc[df["ID"] == sel_id, "CLIENTE"].iloc[0])
                datos["Concepto (Factura)"] = match_concepto(cliente_reg, datos["_Conceptos"])
                if len(datos["_Conceptos"]) > 1:
                    st.caption(f"La factura trae {len(datos['_Conceptos'])} descripciones; "
                               f"se tomó la que coincide con «{cliente_reg}».")
                st.success("Datos extraídos de la factura:")
                st.write({k: datos[k] for k in ["Datos (Factura)", "Concepto (Factura)",
                                                "Número Factura", "Fecha Emisión", "Método de Pago"]})
                st.caption(f"Total en el XML: {fmt_money(datos['_Total'])} · UUID: {datos['_UUID']}")
                try:
                    if datos["_Total"] and abs(float(datos["_Total"]) - float(df.loc[df["ID"] == sel_id, "TOTAL"].iloc[0])) > 0.5:
                        st.warning("⚠️ El total del XML no coincide con el TOTAL del registro.")
                except Exception:
                    pass
                if st.button("💾 Guardar factura en el registro"):
                    full = leer_registros()
                    idx = full.index[full["ID"] == sel_id]
                    for col in FACTURA_COLS:
                        full.loc[idx, col] = datos[col]
                    full.loc[idx, "FECHA DE FACTURACION"] = datos["Fecha Emisión"]
                    sobrescribir_registros(full)
                    st.success(f"✅ Factura {datos['Número Factura']} ligada al registro {sel_id}.")
                    st.rerun()

    st.divider()
    master = layout_a_master(df)
    st.metric("Movimientos en el Registro Máster", len(master))
    if master.empty:
        st.info("Aún no hay movimientos en el LayOut.")
    else:
        vista_master = master[MASTER_FACT_COLS].copy()
        for col in ["Volumen", "Subtotal", "IVA", "Total"]:
            vista_master[col] = vista_master[col].map(fmt_money)
        st.dataframe(vista_master, use_container_width=True, hide_index=True)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            master[MASTER_FACT_COLS].to_excel(w, index=False, sheet_name="Registro Maestro")
        st.download_button("⬇️ Descargar Registro Máster en Excel", buf.getvalue(),
                           file_name="registro_master.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# USUARIOS Y ROLES 
credentials = {
    "usernames": {
        "administracion": {
            "name": "Administrador",
            "password": "$2b$12$zCbFijEUg7dddLGJUGTnOudxIgcbp0CISe3ofbgMVEwVH2op/wfqO",
            "role": "admin", "asesor": None,
        },
        "comercial": {
            "name": "Área Comercial",
            "password": "$2b$12$zCbFijEUg7dddLGJUGTnOudxIgcbp0CISe3ofbgMVEwVH2op/wfqO",
            "role": "movimientos", "asesor": None,
        },
    }
}

authenticator = stauth.Authenticate(
    credentials,
    cookie_name="facturacion_finarq2",
    key="clave_secreta_finarq_2026",
    cookie_expiry_days=1,
)


# CAPA DE DATOS  (Google Sheets con fallback a CSV local)
CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "layout.csv")

CLIENTES_FX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "clientes_fx.txt")


def cargar_clientes_fx() -> list:
    if os.path.exists(CLIENTES_FX_PATH):
        with open(CLIENTES_FX_PATH, encoding="utf-8") as f:
            xs = [l.strip() for l in f if l.strip()]
        return xs or list(CLIENTES_FX_DEFAULT)
    return list(CLIENTES_FX_DEFAULT)


def agregar_cliente_fx(nombre: str):
    nombre = nombre.strip()
    xs = cargar_clientes_fx()
    if nombre and nombre not in xs:
        xs.append(nombre)
        os.makedirs(os.path.dirname(CLIENTES_FX_PATH), exist_ok=True)
        with open(CLIENTES_FX_PATH, "w", encoding="utf-8") as f:
            f.write("\n".join(xs))


def _usar_sheets() -> bool:
    try:
        return "gcp_service_account" in st.secrets and "sheet_url" in st.secrets
    except Exception:
        return False


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

        st.write("Correo usado:", creds.service_account_email)
        st.write("Proyecto:", creds.project_id)
        st.write("Sheet ID:", st.secrets["sheet_id"])

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
        ws.update("A1", [COLUMNS])

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
    return df
       


def guardar_registro(fila: dict):
    """Agrega un registro nuevo al backend (Sheets o CSV)."""
    row = [str(fila.get(c, "")).upper() if c in COLUMNS else str(fila.get(c, ""))
       for c in COLUMNS]
    if _usar_sheets():
        _get_worksheet().append_row(row, value_input_option="USER_ENTERED")
    else:
        os.makedirs(os.path.dirname(CSV_PATH), exist_ok=True)
        df = leer_registros()
        df.loc[len(df)] = row
        df.to_csv(CSV_PATH, index=False)
    leer_registros.clear()


def sobrescribir_registros(df: pd.DataFrame):
    df = df.reindex(columns=COLUMNS).fillna("")
    if _usar_sheets():
        ws = _get_worksheet()
        ws.clear()
        ws.update([COLUMNS] + df.astype(str).values.tolist(), "A1")
    else:
        os.makedirs(os.path.dirname(CSV_PATH), exist_ok=True)
        df.to_csv(CSV_PATH, index=False)
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
    return df.apply(lambda r: "🟢 Completo" if es_completo(r) else "🔴 Incompleto", axis=1)

# LOGIN
authenticator.login("Iniciar sesión", location="main")

if not st.session_state.get("authentication_status"):
    st.title("🧾 Herramienta de Facturación FIN.ARQ")
    if st.session_state.get("authentication_status") is False:
        st.error("Usuario o contraseña incorrectos")
    else:
        st.info("Ingresa tus credenciales para continuar.")
    st.stop()

usuario = st.session_state["username"]
rol = credentials["usernames"][usuario]["role"]
nombre = st.session_state["name"]
ES_ADMIN = rol == "admin"

# SIDEBAR
with st.sidebar:
    st.header("⚙️ Sesión")
    st.write(f"👤 {nombre}")
    st.caption(f"Rol: **{rol}**")
    authenticator.logout("Cerrar sesión", location="sidebar")
    st.divider()
    if not _usar_sheets():
        st.warning("Modo local (CSV). Configura Google Sheets en secrets para producción.")
    else:
        st.success("Conectado a Google Sheets ✅")
    if rol in ("admin", "facturacion"):
        if st.button("🔄 Recargar registros"):
            leer_registros.clear()
            st.rerun()

if rol == "movimientos":
    vistas = ["📝 LayOut — Registro"]
else:
    vistas = ["📝 LayOut — Registro", "🏦 Estado de cuenta"]

vista = st.sidebar.radio("Vista", vistas)

st.title("🧾 Facturación FIN.ARQ")

# VISTA 1: LAYOUT — REGISTRO MANUAL

if vista.startswith("📝"):
    st.subheader("LayOut — Registro de movimientos")
 
    df = leer_registros()
 
    # ---------- Formulario ----------
    if ES_ADMIN:
        st.caption("Captura completa. IVA y TOTAL se calculan solos.")
        with st.form("form_admin", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                codigo_sat = st.text_input("CÓDIGO SAT")
                cliente = st.text_input("CLIENTE *")
                concepto = st.text_input("CONCEPTO")
                producto = st.selectbox("PRODUCTO *", PRODUCTOS)
            with c2:
                aliado = st.text_input("Aliado Comercial")
                asesor = st.selectbox("Asesor *", ASESORES)
                base = st.number_input("BASE (Volumen)", min_value=0.0, step=100.0, format="%.2f")
                tasa = st.number_input("Tasa (%) — opcional, autollena el Ingreso", min_value=0.0,
                                       max_value=100.0, step=0.5, format="%.2f")
            with c3:
                subtotal_in = st.number_input("SUBTOTAL (INGRESO)", min_value=0.0, step=100.0, format="%.2f")
                estatus = st.selectbox("ESTATUS DE PAGO", ESTATUS)
            notas = st.text_input("NOTAS")
            enviar = st.form_submit_button("💾 Guardar registro", use_container_width=True)
 
        if enviar:
            if not cliente.strip():
                st.error("El CLIENTE es obligatorio.")
            else:
                subtotal = round(base * tasa / 100, 2) if tasa > 0 else round(subtotal_in, 2)
                iva = round(subtotal * IVA_RATE, 2)
                fila = {c: "" for c in COLUMNS}
                fila.update({
                    "CÓDIGO SAT": codigo_sat.strip(), "CLIENTE": cliente.strip(),
                    "CONCEPTO": concepto.strip(), "PRODUCTO": producto,
                    "Aliado Comercial": aliado.strip(), "Asesor": asesor,
                    "BASE (Volumen)": base, "SUBTOTAL (INGRESO)": subtotal,
                    "IVA": iva, "TOTAL": round(subtotal + iva, 2),
                    "ESTATUS DE PAGO": estatus,
                    "NOTAS": notas.strip(),
                    "ID": siguiente_id(df), "Capturado_Por": nombre,
                })
                guardar_registro(fila)
                st.success(f"✅ Registro {fila['ID']} guardado.")
                df = leer_registros()
    else:
        st.caption("Captura tus movimientos. Administración completará el resto de la información.")
        producto = st.selectbox("PRODUCTO", PRODUCTOS, key="prod_com")
        es_fx = producto == "FX"
        clientes_fx = cargar_clientes_fx()

        with st.form("form_comercial", clear_on_submit=True):
            aliado = st.text_input("Aliado Comercial", value=FX_ALIADO if es_fx else "")
            if es_fx:
                cliente_sel = st.selectbox("CLIENTE *", clientes_fx + ["➕ Otro (agregar)…"])
                cliente_nuevo = st.text_input("Nuevo cliente (solo si elegiste «Otro»)")
            else:
                cliente_sel = st.text_input("CLIENTE *")
                cliente_nuevo = ""
            asesor = st.selectbox("Asesor *", ASESORES)
            concepto = st.text_input("CONCEPTO")
            c4, c5 = st.columns(2)
            with c4:
                base = st.number_input("Volumen", min_value=0.0, step=100.0, format="%.2f")
            with c5:
                tasa = st.number_input("Tasa", min_value=0.0, max_value=100.0, step=0.5, format="%.2f",
                                       help="Subtotal (Ingreso) = Volumen × Tasa / 100")
            codigo_sat = st.text_input("CÓDIGO SAT", value=FX_CODIGO_SAT if es_fx else "")
            enviar = st.form_submit_button("💾 Guardar registro", use_container_width=True)

        if enviar:
            if es_fx and cliente_sel == "➕ Otro (agregar)…":
                cliente = cliente_nuevo.strip()
            else:
                cliente = cliente_sel.strip()
            if not cliente:
                st.error("El CLIENTE es obligatorio.")
            else:
                if es_fx and cliente not in clientes_fx:
                    agregar_cliente_fx(cliente)      
                subtotal = round(base * tasa / 100, 2)
                iva = round(subtotal * IVA_RATE, 2)
                fila = {c: "" for c in COLUMNS}
                fila.update({
                    "CÓDIGO SAT": codigo_sat.strip(), "CLIENTE": cliente,
                    "CONCEPTO": concepto.strip(), "PRODUCTO": producto,
                    "Aliado Comercial": aliado.strip(), "Asesor": asesor,
                    "BASE (Volumen)": base, "SUBTOTAL (INGRESO)": subtotal,
                    "IVA": iva, "TOTAL": round(subtotal + iva, 2),
                    "ESTATUS DE PAGO": "PENDIENTE",
                    "ID": siguiente_id(df), "Capturado_Por": nombre,
                })
                guardar_registro(fila)
                st.success(f"✅ Registro {fila['ID']} guardado (quedará 🔴 hasta que Administración lo complete).")
                df = leer_registros()
 
    st.divider()
    st.subheader("Registros capturados")
 
    if df.empty:
        st.info("Aún no hay registros.")
    else:
        estado = col_estado(df)
        _f = pd.to_datetime(df["FECHA DE FACTURACION"], errors="coerce")
        df_v = df.copy()
        df_v["_Año"] = _f.dt.year
        df_v["_Mes"] = _f.dt.month.map(lambda n: MESES[int(n) - 1] if pd.notna(n) else "")
 
        with st.expander("🔎 Filtros", expanded=False):
            a, b, c = st.columns(3)
            f_est = a.multiselect("Estado", ["🔴 Incompleto", "🟢 Completo"])
            f_cli = b.multiselect("Cliente", sorted([x for x in df_v["CLIENTE"].dropna().unique() if str(x).strip()]))
            f_ase = c.multiselect("Asesor", sorted([x for x in df_v["Asesor"].dropna().unique() if str(x).strip()]))
 
        m = pd.Series(True, index=df_v.index)
        if f_est: m &= estado.isin(f_est)
        if f_cli: m &= df_v["CLIENTE"].isin(f_cli)
        if f_ase: m &= df_v["Asesor"].isin(f_ase)
        filt = df_v[m]
        st.caption(f"Mostrando {len(filt)} de {len(df_v)} registros.")
        es_pag = filt["ESTATUS DE PAGO"].astype(str).str.upper() == "PAGADA"
        filt_pend, filt_pag = filt[~es_pag], filt[es_pag]

        def _tabla_lectura(dfx):
            d = dfx[["CÓDIGO SAT", "CLIENTE", "CONCEPTO", "PRODUCTO", "Asesor",
                     "BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL",
                     "ESTATUS DE PAGO", "ID"]].copy()
            d.insert(0, "Estado", col_estado(dfx))
            for mc in ["BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL"]:
                d[mc] = d[mc].map(fmt_money)
            d = d[["Estado", "ESTATUS DE PAGO"] +
                  [c for c in d.columns if c not in ("Estado", "ESTATUS DE PAGO")]]
            return d

        if ES_ADMIN:
            # ===== PENDIENTES (rejilla editable) =====
            st.markdown(f"#### 🟠 Pendientes ({len(filt_pend)})")
            if filt_pend.empty:
                st.info("No hay registros pendientes.")
            else:
                tabla = filt_pend[LAYOUT_COLS].copy()
                tabla.insert(0, "Estado", col_estado(filt_pend))
                for mc in MONEY_COLS:
                    tabla[mc] = pd.to_numeric(tabla[mc], errors="coerce")
                tabla["🗑️ Eliminar"] = False
                tabla = tabla[["Estado", "ESTATUS DE PAGO"] +
                              [c for c in tabla.columns
                               if c not in ("Estado", "ESTATUS DE PAGO", "🗑️ Eliminar")] +
                              ["🗑️ Eliminar"]]
                cfg = {
                    "Estado": st.column_config.TextColumn("Estado", disabled=True),
                    "ID": st.column_config.TextColumn("ID", disabled=True),
                    "Capturado_Por": st.column_config.TextColumn("Capturado_Por", disabled=True),
                    "IVA": st.column_config.NumberColumn("IVA", disabled=True, format="$%.2f"),
                    "TOTAL": st.column_config.NumberColumn("TOTAL", disabled=True, format="$%.2f"),
                    "BASE (Volumen)": st.column_config.NumberColumn("BASE (Volumen)", format="$%.2f"),
                    "SUBTOTAL (INGRESO)": st.column_config.NumberColumn("SUBTOTAL (INGRESO)", format="$%.2f"),
                    "🗑️ Eliminar": st.column_config.CheckboxColumn("🗑️ Eliminar"),
                }
                editado = st.data_editor(tabla, use_container_width=True, hide_index=True,
                                         num_rows="fixed", key="editor_layout", column_config=cfg)

                ids_borrar = editado.loc[editado["🗑️ Eliminar"] == True, "ID"].tolist()
                col_c, col_e = st.columns([1.4, 1.6])
                with col_c:
                    confirmar = st.checkbox(f"Confirmar borrado ({len(ids_borrar)})", disabled=not ids_borrar)
                with col_e:
                    if st.button("🗑️ Eliminar seleccionados", disabled=not (ids_borrar and confirmar)):
                        full = leer_registros()
                        full = full[~full["ID"].isin(ids_borrar)]
                        sobrescribir_registros(full)
                        st.success(f"🗑️ {len(ids_borrar)} registro(s) eliminado(s).")
                        st.rerun()

                if st.button("💾 Guardar cambios"):
                    full = leer_registros().astype(object)
                    for _, r in editado.iterrows():
                        idx = full.index[full["ID"] == r["ID"]]
                        if len(idx):
                            for col in EDITABLE_COLS:
                                full.loc[idx, col] = r[col]
                    sub = pd.to_numeric(full["SUBTOTAL (INGRESO)"], errors="coerce").fillna(0)
                    full["IVA"] = (sub * IVA_RATE).round(2)
                    full["TOTAL"] = (sub + sub * IVA_RATE).round(2)
                    sobrescribir_registros(full)
                    st.success("✅ Cambios guardados.")
                    st.rerun()

                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    filt_pend[LAYOUT_COLS].to_excel(w, index=False, sheet_name="Pendientes")
                st.download_button("⬇️ Descargar pendientes en Excel", buf.getvalue(),
                                   file_name="layout_pendientes.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.markdown(f"#### ✅ Pagadas ({len(filt_pag)})")
            if filt_pag.empty:
                st.caption("Aún no hay registros pagados.")
            else:
                with st.expander("Ver registros pagados", expanded=False):
                    st.dataframe(_tabla_lectura(filt_pag), use_container_width=True, hide_index=True)
        else:
            st.markdown(f"#### 🟠 Pendientes ({len(filt_pend)})")
            if filt_pend.empty:
                st.info("No hay registros pendientes.")
            else:
                st.dataframe(_tabla_lectura(filt_pend), use_container_width=True, hide_index=True)
            st.markdown(f"#### ✅ Pagadas ({len(filt_pag)})")
            if filt_pag.empty:
                st.caption("Aún no hay registros pagados.")
            else:
                st.dataframe(_tabla_lectura(filt_pag), use_container_width=True, hide_index=True)
    if ES_ADMIN:
        st.divider()
        render_facturacion(df)
 


# VISTA 2: FACTURACIÓN (placeholder Fase 2)

else:
    st.subheader("🏦 Estado de cuenta")
    st.caption("Sube el PDF del estado de cuenta.")

    df = leer_registros()
    edc = st.file_uploader("PDF del estado de cuenta", type=["pdf"])

    if edc is not None and not df.empty:
        try:
            deps = parse_estado_cuenta(edc.getvalue())
        except Exception as ex:
            st.error(f"No se pudo leer el estado de cuenta: {ex}")
            deps = []
        st.caption(f"Depósitos detectados: {len(deps)} · "
                   f"total {fmt_money(sum(d['monto'] for d in deps))}")

        filas = []
        for _, r in df.iterrows():
            dep = match_deposito(r["TOTAL"], deps)
            if dep:
                filas.append({
                    "Aplicar": True,
                    "ID": r["ID"], "Cliente": r["CLIENTE"],
                    "Total": fmt_money(r["TOTAL"]), "Pago Real": dep["fecha"],
                    "Comprobante de Pago": dep["concepto"],
                    "_pago_real": dep["fecha"], "_comprobante": dep["concepto"],
                })
        if not filas:
            st.warning("Ningún registro coincidió por monto con los depósitos del estado de cuenta.")
        else:
            prev = pd.DataFrame(filas)
            st.markdown("**Conciliación encontrada** (revisa y marca cuáles aplicar):")
            editado = st.data_editor(
                prev[["Aplicar", "ID", "Cliente", "Total", "Pago Real",
                       "Comprobante de Pago"]],
                use_container_width=True, hide_index=True, num_rows="fixed",
                key="editor_edc",
                column_config={c: st.column_config.TextColumn(c, disabled=True)
                               for c in ["ID", "Cliente", "Total", "Pago Real",
                                          "Comprobante de Pago"]},
            )
            if st.button("💾 Aplicar pagos a los registros"):
                aplicar = editado[editado["Aplicar"] == True]["ID"].tolist()
                if aplicar:
                    full = leer_registros()
                    ref = prev.set_index("ID")
                    for rid in aplicar:
                        idx = full.index[full["ID"] == rid]
                        full.loc[idx, "Pago Real"] = ref.loc[rid, "_pago_real"]
                        full.loc[idx, "Comprobante de Pago"] = ref.loc[rid, "_comprobante"]
                        full.loc[idx, "ESTATUS DE PAGO"] = "PAGADA"
                    sobrescribir_registros(full)
                    st.success(f"✅ Se aplicó el pago a {len(aplicar)} registro(s).")
                    st.rerun()
                else:
                    st.info("No marcaste ningún registro.")
    elif df.empty:
        st.info("Aún no hay registros en el LayOut.")
    if not df.empty:
        st.divider()
        st.markdown("### 📋 Registro Máster completo")
        master = layout_a_master(df)
        vm = master.copy()
        for col in ["Volumen", "Subtotal", "IVA", "Total"]:
            vm[col] = vm[col].map(fmt_money)
        st.dataframe(vm, use_container_width=True, hide_index=True)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            master.to_excel(w, index=False, sheet_name="Registro Maestro")
        st.download_button("⬇️ Descargar Registro Máster completo", buf.getvalue(),
                           file_name="registro_master_completo.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
