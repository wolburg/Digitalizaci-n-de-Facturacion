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



def aplicar_factura(full: pd.DataFrame, sel_id: str, datos: dict):
    full = full.astype(object)
    idx = full.index[full["ID"] == sel_id]
    if len(idx) == 0:
        raise ValueError(f"No se encontró el registro {sel_id}")

    for col in FACTURA_COLS:
        full.loc[idx, col] = datos[col]
    full.loc[idx, "FECHA DE FACTURACION"] = datos["Fecha Emisión"]
    full.loc[idx, "FOLIO (LISTO P/FACTURA)"] = datos["Número Factura"]

    corregido = False
    m = montos_xml(datos)
    reg_total = pd.to_numeric(full.loc[idx[0], "TOTAL"], errors="coerce")
    if m is not None and (pd.isna(reg_total) or abs(m[2] - float(reg_total)) > 0.01):
        full.loc[idx, "SUBTOTAL (INGRESO)"] = m[0]
        full.loc[idx, "IVA"] = m[1]
        full.loc[idx, "TOTAL"] = m[2]
        corregido = True
    return full, corregido

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
                full, corregido = aplicar_factura(leer_registros(), sel_id, datos)
                sobrescribir_registros(full)
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
                    sobrescribir_registros(pd.concat([full, pd.DataFrame(filas)], ignore_index=True))
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
    return df.apply(lambda r: "🟢" if es_completo(r) else "🔴", axis=1)

def _color_fila(row):
    est = str(row.get("ESTATUS DE PAGO", "")).upper()
    if est == "LISTO":
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
    st.subheader("Registros capturados")
 
    if df.empty:
        st.info("Aún no hay registros.")
    else:
        estado = col_estado(df)
        _f = pd.to_datetime(df["FECHA DE FACTURACION"], errors="coerce")
        _fp = df["FECHA DE PAGO"].fillna("").astype(str).str.strip()
        _pr = df["Pago Real"].fillna("").astype(str)
        _p = pd.to_datetime(_fp.where(_fp != "", _pr), errors="coerce")

        df_v = df.copy()
        df_v["_Año"] = _f.dt.year
        df_v["_Mes"] = _f.map(lambda d: f"{MESES_CORTO[d.month - 1]}-{d.year}" if pd.notna(d) else "")
        df_v["_FFact"] = _f
        df_v["_FPago"] = _p

        def _opts(col):
            return sorted([x for x in df_v[col].dropna().unique() if str(x).strip()])

        st.markdown("##### 🔎 Filtros")
        r1 = st.columns(4)
        f_est = r1[0].multiselect("Estatus de pago", ESTATUS)
        f_cli = r1[1].multiselect("Cliente", _opts("CLIENTE"))
        f_ase = r1[2].multiselect("Asesor", _opts("Asesor"))
        f_pro = r1[3].multiselect("Producto", _opts("PRODUCTO"))
        r2 = st.columns(3)
        f_ali = r2[0].multiselect("Aliado comercial", _opts("Aliado Comercial"))
        f_ffa = r2[1].date_input("Fecha de facturación (rango)", value=(), format="YYYY-MM-DD", key="f_ffa")
        f_fpa = r2[2].date_input("Fecha de pago (rango)", value=(), format="YYYY-MM-DD", key="f_fpa")

        m = pd.Series(True, index=df_v.index)
        if f_est: m &= df_v["ESTATUS DE PAGO"].astype(str).str.upper().isin(f_est)
        if f_cli: m &= df_v["CLIENTE"].isin(f_cli)
        if f_ase: m &= df_v["Asesor"].isin(f_ase)
        if f_pro: m &= df_v["PRODUCTO"].isin(f_pro)
        if f_ali: m &= df_v["Aliado Comercial"].isin(f_ali)
        if len(f_ffa) == 2:
            m &= df_v["_FFact"].between(pd.Timestamp(f_ffa[0]), pd.Timestamp(f_ffa[1]))
        if len(f_fpa) == 2:
            m &= df_v["_FPago"].between(pd.Timestamp(f_fpa[0]), pd.Timestamp(f_fpa[1]))
        es_pag = df_v["ESTATUS DE PAGO"].astype(str).str.upper() == "PAGADA"
        filt_pend = df_v[m & ~es_pag]
        todas_pag = df[es_pag]

        def _tabla_lectura(dfx, fecha_pago=False):
            cols = ["CÓDIGO SAT", "CLIENTE", "CONCEPTO", "PRODUCTO", "Asesor",
                    "BASE (Volumen)", "SUBTOTAL (INGRESO)", "IVA", "TOTAL",
                    "ESTATUS DE PAGO", "ID"]
            if fecha_pago:
                cols.insert(cols.index("ESTATUS DE PAGO"), "FECHA DE PAGO")
            d = dfx[cols].copy()
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
                    "ESTATUS DE PAGO": st.column_config.SelectboxColumn("ESTATUS DE PAGO", options=ESTATUS),
                }
                editado = st.data_editor(tabla.style.apply(_color_fila, axis=1),
                                        use_container_width=True, hide_index=True,
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
                            full.loc[idx, "Pago Real"] = r["FECHA DE PAGO"]
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

            st.markdown("#### ✅ Pagadas")
            _flash_pag = st.session_state.pop("flash_pag", None)
            if _flash_pag:
                getattr(st, _flash_pag[0])(_flash_pag[1])
            if todas_pag.empty:
                st.caption("Aún no hay registros pagados.")
            else:
                filt_pag = filtros_master(todas_pag, "pag", estatus=False)
                st.caption(f"Mostrando {len(filt_pag)} de {len(todas_pag)} registros pagados.")
                with st.expander("Ver registros pagados", expanded=False):
                    tabla_pag = _tabla_lectura(filt_pag, fecha_pago=True)
                    tabla_pag["FECHA DE PAGO"] = pd.to_datetime(tabla_pag["FECHA DE PAGO"], errors="coerce").dt.date
                    cfg_pag = {c: st.column_config.TextColumn(c, disabled=True)
                               for c in tabla_pag.columns if c != "FECHA DE PAGO"}
                    cfg_pag["FECHA DE PAGO"] = st.column_config.DateColumn(
                        "FECHA DE PAGO", format="YYYY-MM-DD", help="Editable a mano")
                    editado_pag = st.data_editor(
                        tabla_pag.style.apply(_color_fila, axis=1),
                        use_container_width=True, hide_index=True, num_rows="fixed",
                        key=f"editor_pagadas_{abs(hash(tuple(filt_pag['ID'])))}",
                        column_config=cfg_pag)
                    if st.button("💾 Guardar fechas de pago"):
                        full = leer_registros().astype(object)
                        cambiados = []
                        for _, r in editado_pag.iterrows():
                            nueva = _fecha_iso(r["FECHA DE PAGO"])
                            idx = full.index[full["ID"] == r["ID"]]
                            if len(idx) and _fecha_iso(full.loc[idx[0], "FECHA DE PAGO"]) != nueva:
                                full.loc[idx, "FECHA DE PAGO"] = nueva
                                full.loc[idx, "Pago Real"] = nueva
                                cambiados.append(r["ID"])
                        if cambiados:
                            sobrescribir_registros(full)
                            st.session_state["flash_pag"] = (
                                "success", f"✅ Fecha de pago actualizada en: {', '.join(cambiados)}.")
                        else:
                            st.session_state["flash_pag"] = ("info", "No hubo cambios en las fechas de pago.")
                        st.rerun()
                etiquetas_pag = {r["ID"]: f'{r["ID"]} — {r["CLIENTE"]} · {fmt_money(r["TOTAL"])}'
                                  for _, r in filt_pag.iterrows()}
                ids_borrar_pag = st.multiselect(
                    "Eliminar registro(s) pagado(s)",
                    options=list(etiquetas_pag.keys()),
                    format_func=lambda x: etiquetas_pag[x],
                    key="borrar_pag_sel",
                )
                cpg, epg = st.columns([1.4, 1.6])
                with cpg:
                    confirmar_pag = st.checkbox(f"Confirmar borrado ({len(ids_borrar_pag)})",
                                                 disabled=not ids_borrar_pag, key="confirmar_pag")
                with epg:
                    if st.button("🗑️ Eliminar pagados seleccionados",
                                  disabled=not (ids_borrar_pag and confirmar_pag)):
                        full = leer_registros()
                        full = full[~full["ID"].isin(ids_borrar_pag)]
                        sobrescribir_registros(full)
                        st.success(f"🗑️ {len(ids_borrar_pag)} registro(s) eliminado(s).")
                        st.rerun()
        else:
            st.markdown(f"#### 🟠 Pendientes ({len(filt_pend)})")
            if filt_pend.empty:
                st.info("No hay registros pendientes.")
            else:
                st.dataframe(_tabla_lectura(filt_pend).style.apply(_color_fila, axis=1), use_container_width=True, hide_index=True)
            st.markdown("#### ✅ Pagadas")
            if todas_pag.empty:
                st.caption("Aún no hay registros pagados.")
            else:
                filt_pag = filtros_master(todas_pag, "pag", estatus=False)
                st.caption(f"Mostrando {len(filt_pag)} de {len(todas_pag)} registros pagados.")
                st.dataframe(_tabla_lectura(filt_pag, fecha_pago=True).style.apply(_color_fila, axis=1),use_container_width=True, hide_index=True)
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
                    full = leer_registros().astype(object)
                    ref = prev.set_index("ID")
                    corregidos = []
                    for rid in aplicar:
                        idx = full.index[full["ID"] == rid]
                        if len(idx) == 0:
                            continue
                        pago_real = ref.loc[rid, "_pago_real"]
                        if _fecha_iso(full.loc[idx[0], "FECHA DE PAGO"]) != pago_real:
                            full.loc[idx, "FECHA DE PAGO"] = pago_real
                            corregidos.append(rid)
                        full.loc[idx, "Pago Real"] = pago_real
                        full.loc[idx, "Comprobante de Pago"] = ref.loc[rid, "_comprobante"]
                        full.loc[idx, "ESTATUS DE PAGO"] = "PAGADA"
                    sobrescribir_registros(full)
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
