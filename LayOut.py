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


# ----------------------------------------------------------------------------
# CONFIG GENERAL
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Facturación FIN.ARQ — LayOut", page_icon="🧾", layout="wide")

IVA_RATE = 0.16

ASESORES = ["Jorge Orozco", "Andrea Dávila", "Ximena Mora",
            "Marco Ochoa", "Eduardo Jáquez", "Negocios-FINARQ"]

PRODUCTOS = ["FX", "Terminal Punto de Venta", "Crédito",
             "Arrendamiento", "Honorarios / Asesoría", "Credito TPV"]

ESTATUS = ["LISTO", "TIMBRADA", "PAGADA"]

# FALTANTE solo se muestra en pantalla (pago programado vencido y no PAGADA); no se guarda
ESTATUS_VISTA = ["FALTANTE"] + ESTATUS
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


def filtros_master(df: pd.DataFrame, key: str, estatus: bool = True) -> pd.DataFrame:
    if df.empty:
        return df
    _f = pd.to_datetime(df["FECHA DE FACTURACION"], errors="coerce")
    _fp = df["FECHA DE PAGO"].fillna("").astype(str).str.strip()
    _pr = df["Pago Real"].fillna("").astype(str)
    _p = pd.to_datetime(_fp.where(_fp != "", _pr), errors="coerce")

    def _opts(col):
        return sorted([x for x in df[col].dropna().unique() if str(x).strip()])

    st.markdown("##### 🔎 Filtros")
    r1 = st.columns(4 if estatus else 3)
    f_est = r1[0].multiselect("Estatus de pago", ESTATUS, key=f"{key}_est") if estatus else []
    f_cli = r1[-3].multiselect("Cliente", _opts("CLIENTE"), key=f"{key}_cli")
    f_ase = r1[-2].multiselect("Asesor", _opts("Asesor"), key=f"{key}_ase")
    f_pro = r1[-1].multiselect("Producto", _opts("PRODUCTO"), key=f"{key}_pro")
    r2 = st.columns(4)
    f_ali = r2[0].multiselect("Aliado comercial", _opts("Aliado Comercial"), key=f"{key}_ali")
    f_met = r2[1].multiselect("Método de pago", _opts("Método de Pago"), key=f"{key}_met")
    f_ffa = r2[2].date_input("Fecha de facturación (rango)", value=(), format="YYYY-MM-DD", key=f"{key}_ffa")
    f_fpa = r2[3].date_input("Fecha de pago (rango)", value=(), format="YYYY-MM-DD", key=f"{key}_fpa")

    m = pd.Series(True, index=df.index)
    if f_est: m &= df["ESTATUS DE PAGO"].astype(str).str.upper().isin(f_est)
    if f_cli: m &= df["CLIENTE"].isin(f_cli)
    if f_ase: m &= df["Asesor"].isin(f_ase)
    if f_pro: m &= df["PRODUCTO"].isin(f_pro)
    if f_ali: m &= df["Aliado Comercial"].isin(f_ali)
    if f_met: m &= df["Método de Pago"].isin(f_met)
    if len(f_ffa) == 2:
        m &= _f.between(pd.Timestamp(f_ffa[0]), pd.Timestamp(f_ffa[1]))
    if len(f_fpa) == 2:
        m &= _p.between(pd.Timestamp(f_fpa[0]), pd.Timestamp(f_fpa[1]))
    return df[m]


def render_facturacion(df: pd.DataFrame):
    st.header("📄 Facturación")
    st.markdown("### 📎 Adjuntar factura (XML CFDI) a un registro")

    _flash = st.session_state.pop("flash", None)
    if _flash:
        getattr(st, _flash[0])(_flash[1])

    sel_id = st.session_state.get("factura_sel_id")
    if sel_id not in df["ID"].values:
        sel_id = None

    xml_file = None
    if df.empty:
        st.info("Primero captura movimientos en el LayOut.")
    elif not sel_id:
        st.caption("Marca el 📎 de la fila correspondiente, hasta la derecha en la tabla de abajo, para adjuntar su factura.")
    else:
        cliente_reg_lbl = df.loc[df["ID"] == sel_id, "CLIENTE"].iloc[0]
        cab, cbb = st.columns([3, 1])
        cab.success(f"Registro seleccionado: **{sel_id}** — {cliente_reg_lbl}")
        if cbb.button("✖️ Cambiar registro"):
            st.session_state["factura_sel_id"] = None
            st.rerun()
        xml_file = st.file_uploader("XML de la factura", type=["xml"], key=f"xmlup_{sel_id}")

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
                    st.warning("⚠️ El total del XML no coincide con el TOTAL del registro. Se corregirá al guardar.")
            except Exception:
                pass
            if st.button("💾 Guardar factura en el registro"):
                cambios, corregido = aplicar_factura(df.loc[df["ID"] == sel_id].iloc[0], datos)
                actualizar_registros({sel_id: cambios})
                if corregido:
                    st.session_state["flash"] = (
                        "warning",
                        f"Factura {datos['Número Factura']} ligada a {sel_id}. "
                        f"El total no coincidía: SUBTOTAL, IVA y TOTAL se corrigieron con el XML "
                        f"(Total {fmt_money(datos['_Total'])}).")
                else:
                    st.session_state["flash"] = (
                        "success", f"Factura {datos['Número Factura']} ligada a {sel_id}.")
                st.session_state["factura_sel_id"] = None
                st.rerun()

    st.divider()
    st.markdown("### 📦 Carga masiva de facturas Kushki (XML)")
    st.caption(f"Sube un ZIP o varios XML. Cada factura crea un registro nuevo con "
               f"Producto «{KUSHKI_PRODUCTO}», Aliado «{KUSHKI_ALIADO}» y Asesor «{KUSHKI_ASESOR}». "
               f"Solo el Volumen se completa a mano.")
    lote = st.file_uploader("ZIP o XMLs", type=["zip", "xml"], accept_multiple_files=True,
                            key=f"lote_xml_{st.session_state.get('lote_n', 0)}")
    if lote:
        existentes = set(df["Número Factura"].astype(str).str.upper().str.strip()) - {""}
        nuevas, omitidas = [], []
        try:
            xmls = extraer_xmls(lote)
        except zipfile.BadZipFile:
            xmls = []
            st.error("El ZIP está dañado o no es un ZIP válido.")
        for nom, contenido in xmls:
            try:
                datos = parse_cfdi(contenido)
            except Exception:
                omitidas.append(f"{nom}: no se pudo leer")
                continue
            if datos["_Tipo"] and datos["_Tipo"] != "I":
                omitidas.append(f"{nom}: no es factura de ingreso (tipo {datos['_Tipo']})")
                continue
            num = datos["Número Factura"].upper().strip()
            if num and num in existentes:
                omitidas.append(f"{nom}: la factura {datos['Número Factura']} ya está en el registro")
                continue
            existentes.add(num)
            nuevas.append((nom, datos))

        if omitidas:
            with st.expander(f"⏭️ {len(omitidas)} archivo(s) omitido(s)"):
                st.markdown("\n".join(f"- {o}" for o in omitidas))
        if not nuevas:
            st.info("No hay facturas nuevas para crear.")
        else:
            prev_lote = pd.DataFrame([{
                "Aplicar": True, "Archivo": nom, "Factura": d["Número Factura"],
                "Cliente": d["_Nombre"], "Concepto": d["Concepto (Factura)"],
                "Emisión": d["Fecha Emisión"], "Método": d["Método de Pago"],
                "Total": fmt_money(d["_Total"]),
                "BASE (Volumen)": float("nan"),
            } for nom, d in nuevas])
            ro = ["Archivo", "Factura", "Cliente", "Concepto", "Emisión", "Método", "Total"]
            cfg_lote = {c: st.column_config.TextColumn(c, disabled=True) for c in ro}
            cfg_lote["BASE (Volumen)"] = st.column_config.NumberColumn(
                "BASE (Volumen)", min_value=0.0, format="$%.2f")
            st.markdown(f"**{len(nuevas)} factura(s) nuevas** — captura el Volumen (o déjalo para después) y crea los registros:")
            editado_lote = st.data_editor(
                prev_lote, use_container_width=True, hide_index=True, num_rows="fixed",
                key=f"editor_lote_{abs(hash(tuple(n for n, _ in nuevas)))}", column_config=cfg_lote)
            if st.button("➕ Crear registros desde los XML"):
                full = leer_registros()
                n_id = int(siguiente_id(full).replace("MOV-", ""))
                filas = []
                for i, r in editado_lote.iterrows():
                    if r["Aplicar"] != True:
                        continue
                    fila = fila_desde_xml(nuevas[i][1], f"MOV-{n_id:05d}", nombre)
                    n_id += 1
                    if pd.notna(r["BASE (Volumen)"]):
                        fila["BASE (Volumen)"] = r["BASE (Volumen)"]
                    filas.append(fila)
                if filas:
                    agregar_registros(filas)
                    st.session_state["flash"] = (
                        "success", f"✅ {len(filas)} registro(s) creados desde XML "
                                   f"({filas[0]['ID']} a {filas[-1]['ID']}).")
                    st.session_state["lote_n"] = st.session_state.get("lote_n", 0) + 1
                    st.rerun()
                else:
                    st.info("No marcaste ninguna factura.")

    st.divider()
    df_m = filtros_master(df, "mf")
    master = layout_a_master(df_m)
    st.metric("Movimientos en el Registro Máster", len(master))
    if master.empty:
        st.info("No hay movimientos para mostrar.")
    else:
        vista_master = master[MASTER_FACT_COLS].copy().reset_index(drop=True)
        for col in ["Volumen", "Subtotal", "IVA", "Total"]:
            vista_master[col] = vista_master[col].map(fmt_money)

        ids_list = df_m["ID"].tolist()
        prev_sel = st.session_state.get("factura_sel_id")
        vista_master_sel = vista_master.copy()
        vista_master_sel["📎"] = [rid == prev_sel for rid in ids_list[:len(vista_master_sel)]]

        cfg_master = {c: st.column_config.TextColumn(c, disabled=True) for c in MASTER_FACT_COLS}
        cfg_master["📎"] = st.column_config.CheckboxColumn("📎", help="Marca para adjuntar factura a este registro")

        editado_master = st.data_editor(
            vista_master_sel, use_container_width=True, hide_index=True,
            key=f"editor_master_pick_{abs(hash(tuple(ids_list)))}", column_config=cfg_master,
        )

        marcados_idx = editado_master.index[editado_master["📎"] == True].tolist()
        marcados_ids = [ids_list[i] for i in marcados_idx if i < len(ids_list)]
        nuevos = [x for x in marcados_ids if x != prev_sel]
        if nuevos:
            st.session_state["factura_sel_id"] = nuevos[0]
            st.rerun()
        elif not marcados_ids and prev_sel:
            st.session_state["factura_sel_id"] = None
            st.rerun()

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
    """ESTATUS DE PAGO con FALTANTE cuando el pago programado ya venció y no está PAGADA."""
    est = df["ESTATUS DE PAGO"].astype(str).str.upper()
    pp = pago_programado(df)
    vencido = (pp != "") & (pp < date.today().isoformat()) & (est != "PAGADA")
    return est.where(~vencido, "FALTANTE")


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
    if est == "FALTANTE":
        css = "background-color: rgba(239, 68, 68, 0.30)"
    elif est == "LISTO":
        css = "background-color: rgba(59, 130, 246, 0.30)"   
    elif est == "PAGADA":
        css = "background-color: rgba(34, 197, 94, 0.30)"    
    else:
        css = ""                                             
    return [css] * len(row)


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
                    "ESTATUS DE PAGO": "LISTO",
                    "ID": siguiente_id(df), "Capturado_Por": nombre,
                })
                guardar_registro(fila)
                st.success(f"✅ Registro {fila['ID']} guardado (quedará 🔴 hasta que Administración lo complete).")
                df = leer_registros()
 
    st.divider()
    st.subheader("📋 Registros")
    _flash = st.session_state.pop("flash_layout", None)
    if _flash:
        getattr(st, _flash[0])(_flash[1])

    if df.empty:
        st.info("Aún no hay registros.")
    else:
        _f = pd.to_datetime(df["FECHA DE FACTURACION"], errors="coerce")
        _fp = df["FECHA DE PAGO"].fillna("").astype(str).str.strip()
        _pr = df["Pago Real"].fillna("").astype(str)
        _p = pd.to_datetime(_fp.where(_fp != "", _pr), errors="coerce")

        df_v = df.copy()
        df_v["ESTATUS DE PAGO"] = estatus_vista(df)
        df_v["PAGO PROGRAMADO"] = pago_programado(df)
        df_v["_FFact"] = _f
        df_v["_FPago"] = _p

        def _opts(col):
            return sorted([x for x in df_v[col].dropna().unique() if str(x).strip()])

        st.markdown("##### 🔎 Filtros")
        r1 = st.columns(4)
        f_est = r1[0].multiselect("Estatus de pago", ESTATUS_VISTA)
        f_cli = r1[1].multiselect("Cliente", _opts("CLIENTE"))
        f_ase = r1[2].multiselect("Asesor", _opts("Asesor"))
        f_pro = r1[3].multiselect("Producto", _opts("PRODUCTO"))
        r2 = st.columns(3)
        f_ali = r2[0].multiselect("Aliado comercial", _opts("Aliado Comercial"))
        f_ffa = r2[1].date_input("Fecha de facturación (rango)", value=(), format="YYYY-MM-DD", key="f_ffa")
        f_fpa = r2[2].date_input("Fecha de pago (rango)", value=(), format="YYYY-MM-DD", key="f_fpa")

        m = pd.Series(True, index=df_v.index)
        if f_est: m &= df_v["ESTATUS DE PAGO"].isin(f_est)
        if f_cli: m &= df_v["CLIENTE"].isin(f_cli)
        if f_ase: m &= df_v["Asesor"].isin(f_ase)
        if f_pro: m &= df_v["PRODUCTO"].isin(f_pro)
        if f_ali: m &= df_v["Aliado Comercial"].isin(f_ali)
        if len(f_ffa) == 2:
            m &= df_v["_FFact"].between(pd.Timestamp(f_ffa[0]), pd.Timestamp(f_ffa[1]))
        if len(f_fpa) == 2:
            m &= df_v["_FPago"].between(pd.Timestamp(f_fpa[0]), pd.Timestamp(f_fpa[1]))
        filt = df_v[m]

        # Columnas de la tabla: PAGO PROGRAMADO (solo lectura) junto a FECHA DE PAGO
        cols_tabla = [c for c in LAYOUT_COLS if c != "ESTATUS DE PAGO"]
        cols_tabla.insert(cols_tabla.index("FECHA DE PAGO") + 1, "PAGO PROGRAMADO")

        def _tabla_lectura(dfx):
            cols = ["CÓDIGO SAT", "CLIENTE", "CONCEPTO", "PRODUCTO", "Asesor",
                    "BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL",
                    "FECHA DE FACTURACION", "FECHA DE PAGO", "PAGO PROGRAMADO", "ID"]
            d = dfx[cols].copy()
            d.insert(0, "Estado", col_estado(dfx))
            d.insert(1, "ESTATUS DE PAGO", dfx["ESTATUS DE PAGO"])
            for mc in ["BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL"]:
                d[mc] = d[mc].map(fmt_money)
            return agrupar_por_mes(d)

        st.markdown(f"#### Registros ({len(filt)})")
        st.caption("Ordenados por mes de facturación: FALTANTE, LISTO, TIMBRADA y PAGADA. "
                   "FALTANTE = ya pasó el pago programado y no está pagada.")
        if filt.empty:
            st.info("No hay registros con esos filtros.")
        elif ES_ADMIN:
            # Cambios sin guardar, por ID: sobreviven a los cambios de filtro
            pend = st.session_state.setdefault("layout_pend", {})
            ver = st.session_state.get("layout_ver", 0)

            orig = filt[cols_tabla].copy()
            orig.insert(0, "Estado", col_estado(filt))
            orig.insert(1, "ESTATUS DE PAGO", filt["ESTATUS DE PAGO"])
            for mc in MONEY_COLS:
                orig[mc] = pd.to_numeric(orig[mc], errors="coerce")
            orig_id = orig.drop_duplicates("ID").set_index("ID")
            firma = abs(hash((ver, int(pd.util.hash_pandas_object(orig, index=False).sum()))))

            def _pendiente(col, v):
                """Normaliza un valor editado; None si no se debe guardar."""
                if col not in EDITABLE_COLS or (col == "ESTATUS DE PAGO" and v == "FALTANTE"):
                    return None
                return v.upper() if col in UPPER_COLS and isinstance(v, str) else v

            # La tabla que recibe el editor solo se rearma cuando cambian los filtros o los datos;
            # si cambiara en cada edición, Streamlit reiniciaría el editor y se perderían cambios.
            snap = st.session_state.get("layout_snap")
            if not snap or snap[0] != firma:
                if snap:
                    # Rescatar lo editado en la tabla anterior que aún no se había procesado
                    previo = st.session_state.get(f"editor_layout_{snap[0]}") or {}
                    for fila, cols in previo.get("edited_rows", {}).items():
                        rid = snap[1].iloc[int(fila)]["ID"]
                        for col, v in cols.items():
                            v = _pendiente(col, v)
                            if rid and v is not None:
                                pend.setdefault(rid, {})[col] = v
                base = orig.copy()
                for rid, cols in pend.items():
                    idx = base.index[base["ID"] == rid]
                    for col, v in cols.items():
                        base.loc[idx, col] = v
                base["🗑️ Eliminar"] = False
                snap = (firma, agrupar_por_mes(base))
                st.session_state["layout_snap"] = snap
            tabla = snap[1]
            cfg = {
                "Estado": st.column_config.TextColumn("Estado", disabled=True),
                "ID": st.column_config.TextColumn("ID", disabled=True),
                "Capturado_Por": st.column_config.TextColumn("Capturado_Por", disabled=True),
                "PAGO PROGRAMADO": st.column_config.TextColumn("PAGO PROGRAMADO", disabled=True),
                "IVA": st.column_config.NumberColumn("IVA", disabled=True, format="$%.2f"),
                "TOTAL": st.column_config.NumberColumn("TOTAL", disabled=True, format="$%.2f"),
                "BASE (Volumen)": st.column_config.NumberColumn("BASE (Volumen)", format="$%.2f"),
                "SUBTOTAL (INGRESO)": st.column_config.NumberColumn("SUBTOTAL (INGRESO)", format="$%.2f"),
                "🗑️ Eliminar": st.column_config.CheckboxColumn("🗑️ Eliminar"),
                "ESTATUS DE PAGO": st.column_config.SelectboxColumn("ESTATUS DE PAGO", options=ESTATUS_VISTA),
            }
            editado = st.data_editor(tabla.style.apply(_color_fila, axis=1),
                                    use_container_width=True, hide_index=True, num_rows="fixed",
                                    key=f"editor_layout_{firma}", column_config=cfg)
            # Las filas en blanco (separadores de mes) no tienen ID y se ignoran
            editado = editado[editado["ID"].fillna("").astype(str).str.strip() != ""]

            # Pendientes de las filas visibles = lo que difiere de lo guardado
            for _, r in editado.iterrows():
                rid = r["ID"]
                cambios = {}
                for col in EDITABLE_COLS:
                    v = _pendiente(col, r[col])
                    if v is not None and not _igual(v, orig_id.at[rid, col]):
                        cambios[col] = v
                if cambios:
                    pend[rid] = cambios
                else:
                    pend.pop(rid, None)

            if pend:
                ocultos = len(set(pend) - set(orig["ID"]))
                st.warning(f"✏️ Tienes cambios sin guardar en {len(pend)} registro(s)"
                           + (f" ({ocultos} no se ven con los filtros actuales)" if ocultos else "")
                           + ". Se conservan aunque cambies los filtros; dale «💾 Guardar cambios» para guardarlos.")

            ids_borrar = editado.loc[editado["🗑️ Eliminar"] == True, "ID"].tolist()
            col_c, col_e = st.columns([1.4, 1.6])
            with col_c:
                confirmar = st.checkbox(f"Confirmar borrado ({len(ids_borrar)})", disabled=not ids_borrar)
            with col_e:
                if st.button("🗑️ Eliminar seleccionados", disabled=not (ids_borrar and confirmar)):
                    eliminar_registros(ids_borrar)
                    for rid in ids_borrar:
                        pend.pop(rid, None)
                    st.session_state["layout_ver"] = ver + 1
                    st.session_state["flash_layout"] = (
                        "success", f"🗑️ {len(ids_borrar)} registro(s) eliminado(s).")
                    st.rerun()

            col_g, col_d = st.columns([1.4, 1.6])
            with col_g:
                guardar = st.button("💾 Guardar cambios", disabled=not pend)
            with col_d:
                if st.button("↩️ Descartar cambios", disabled=not pend):
                    pend.clear()
                    st.session_state["layout_ver"] = ver + 1
                    st.rerun()
            if guardar:
                cambios = {}
                for rid, cols in pend.items():
                    c = dict(cols)
                    # IVA y TOTAL solo se recalculan si cambió el SUBTOTAL de ese registro
                    if "SUBTOTAL (INGRESO)" in c:
                        sub = pd.to_numeric(c["SUBTOTAL (INGRESO)"], errors="coerce")
                        sub = 0.0 if pd.isna(sub) else float(sub)
                        c["IVA"] = round(sub * IVA_RATE, 2)
                        c["TOTAL"] = round(sub + sub * IVA_RATE, 2)
                    if "FECHA DE PAGO" in c:
                        c["Pago Real"] = c["FECHA DE PAGO"]
                    cambios[rid] = c
                faltan = actualizar_registros(cambios)
                pend.clear()
                st.session_state["layout_ver"] = ver + 1
                if faltan:
                    st.session_state["flash_layout"] = (
                        "warning", f"Cambios guardados, excepto en {', '.join(faltan)}: "
                                   f"ya no existen (alguien más los borró).")
                else:
                    st.session_state["flash_layout"] = (
                        "success", f"✅ Cambios guardados en {len(cambios)} registro(s).")
                st.rerun()

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                agrupar_por_mes(filt[["ESTATUS DE PAGO"] + cols_tabla]).to_excel(
                    w, index=False, sheet_name="Registros")
            st.download_button("⬇️ Descargar registros en Excel", buf.getvalue(),
                               file_name="layout_registros.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.dataframe(_tabla_lectura(filt).style.apply(_color_fila, axis=1),
                         use_container_width=True, hide_index=True)
    if ES_ADMIN:
        st.divider()
        render_facturacion(df)
 


# VISTA 2: FACTURACIÓN (placeholder Fase 2)

else:
    st.subheader("🏦 Estado de cuenta")
    st.caption("Sube el PDF del estado de cuenta.")

    _flash = st.session_state.pop("flash_edc", None)
    if _flash:
        getattr(st, _flash[0])(_flash[1])

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
                fp_reg = _fecha_iso(r["FECHA DE PAGO"])
                filas.append({
                    "Aplicar": True,
                    "ID": r["ID"], "Cliente": r["CLIENTE"],
                    "Total": fmt_money(r["TOTAL"]),
                    "Fecha de pago (registro)": fp_reg,
                    "Fecha (estado de cuenta)": dep["fecha"],
                    "Fecha": "✅" if fp_reg == dep["fecha"] else "⚠️ Se corregirá",
                    "Comprobante de Pago": dep["concepto"],
                    "_pago_real": dep["fecha"], "_comprobante": dep["concepto"],
                })
        if not filas:
            st.warning("Ningún registro coincidió por monto con los depósitos del estado de cuenta.")
        else:
            prev = pd.DataFrame(filas)
            st.markdown("**Conciliación encontrada** (revisa y marca cuáles aplicar):")
            n_dif = int((prev["Fecha"] != "✅").sum())
            if n_dif:
                st.warning(f"⚠️ {n_dif} registro(s) tienen una FECHA DE PAGO distinta (o vacía) "
                           f"a la del estado de cuenta. Se corregirá con la fecha del depósito al aplicar.")
            vis_cols = ["ID", "Cliente", "Total", "Fecha de pago (registro)", "Fecha (estado de cuenta)",
                        "Fecha", "Comprobante de Pago"]
            editado = st.data_editor(
                prev[["Aplicar"] + vis_cols],
                use_container_width=True, hide_index=True, num_rows="fixed",
                key="editor_edc",
                column_config={c: st.column_config.TextColumn(c, disabled=True)
                               for c in vis_cols},
            )
            if st.button("💾 Aplicar pagos a los registros"):
                aplicar = editado[editado["Aplicar"] == True]["ID"].tolist()
                if aplicar:
                    ref = prev.set_index("ID")
                    cambios, corregidos = {}, []
                    for rid in aplicar:
                        pago_real = ref.loc[rid, "_pago_real"]
                        cambios[rid] = {"Pago Real": pago_real,
                                        "Comprobante de Pago": ref.loc[rid, "_comprobante"],
                                        "ESTATUS DE PAGO": "PAGADA"}
                        if ref.loc[rid, "Fecha de pago (registro)"] != pago_real:
                            cambios[rid]["FECHA DE PAGO"] = pago_real
                            corregidos.append(rid)
                    actualizar_registros(cambios)
                    msg = f"✅ Se aplicó el pago a {len(aplicar)} registro(s) con éxito."
                    if corregidos:
                        msg += (f" La fecha de pago no coincidía con el estado de cuenta y se corrigió "
                                f"en: {', '.join(corregidos)}.")
                    st.session_state["flash_edc"] = ("warning" if corregidos else "success", msg)
                    st.rerun()
                else:
                    st.info("No marcaste ningún registro.")
    elif df.empty:
        st.info("Aún no hay registros en el LayOut.")
    if not df.empty:
        st.divider()
        st.markdown("### 📋 Registro Máster completo")
        master = layout_a_master(filtros_master(df, "mc"))
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