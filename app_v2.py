"""Facturación FIN.ARQ — Versión 2 (demo).

Misma lógica de captura, corrección y conexión de datos que LayOut.py (vive en core.py);
aquí cambia la forma de presentar la información: una tabla por mes, editable, con filas
de color según el estatus.

    streamlit run app_v2.py
"""
import io
import zipfile
from datetime import date

import pandas as pd
import streamlit as st

from core import *  # noqa: F401,F403  (constantes, lectura de XML/PDF y capa de datos)
from core import _igual, _usar_sheets, _fecha_iso

st.set_page_config(page_title="Facturación FIN.ARQ", page_icon="🧾", layout="wide")


# ----------------------------------------------------------------------------
# ESTATUS
# Se guardan LISTO / TIMBRADA / PAGADA / CANCELADA. VENCIDA solo se muestra: se calcula
# cuando ya pasó el pago programado y no está PAGADA ni CANCELADA.
# ----------------------------------------------------------------------------
ESTADOS = {
    "VENCIDA":   {"icono": "🔴", "tinte": "rgba(239, 68, 68, 0.30)",
                  "ayuda": "Ya pasó la fecha programada de pago y no se ha cobrado. Se calcula sola."},
    "LISTO":     {"icono": "🔵", "tinte": "rgba(59, 130, 246, 0.30)",
                  "ayuda": "Movimiento capturado, listo para facturar."},
    "TIMBRADA":  {"icono": "🟡", "tinte": "rgba(250, 178, 25, 0.28)",
                  "ayuda": "Factura emitida (timbrada); esperando el pago."},
    "PAGADA":    {"icono": "🟢", "tinte": "rgba(34, 197, 94, 0.30)",
                  "ayuda": "Pago recibido."},
    "CANCELADA": {"icono": "⚫", "tinte": "rgba(120, 120, 120, 0.25)",
                  "ayuda": "Movimiento o factura cancelada. No cuenta en los totales ni se vuelve vencida."},
}


def etiqueta(est: str) -> str:
    e = ESTADOS.get(est)
    return f"{e['icono']} {est}" if e else str(est)


ETIQUETA_A_ESTADO = {etiqueta(e): e for e in ESTATUS_VISTA}

MESES_VISIBLES = 6


# ----------------------------------------------------------------------------
# AYUDANTES
# ----------------------------------------------------------------------------
def a_fecha(v):
    """Texto de la hoja -> date (acepta AAAA-MM-DD y DD/MM/AAAA); None si está vacío."""
    s = str(v if v is not None else "").strip()
    if not s or s.lower() == "nan":
        return None
    if len(s) >= 10 and s[4] == "-":
        f = pd.to_datetime(s[:10], format="%Y-%m-%d", errors="coerce")
    else:
        f = pd.to_datetime(s, dayfirst=True, errors="coerce")
    return f.date() if pd.notna(f) else None


def nombre_mes(p) -> str:
    return f"{MESES_CORTO[p.month - 1]} {p.year}" if pd.notna(p) else "Sin fecha"


def preparar(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega columnas calculadas (prefijo _) para mostrar, filtrar y sumar."""
    d = df.copy()
    if d.empty:
        return d
    d["_est"] = estatus_vista(df)
    d["_pp"] = pd.to_datetime(pago_programado(df), errors="coerce")
    for c, n in [("TOTAL", "_total"), ("SUBTOTAL (INGRESO)", "_sub"), ("IVA", "_iva"),
                 ("BASE (Volumen)", "_base")]:
        d[n] = pd.to_numeric(d[c], errors="coerce")
    d["_ffact"] = pd.to_datetime(d["FECHA DE FACTURACION"].map(a_fecha))
    d["_fpago"] = pd.to_datetime(d["FECHA DE PAGO"].map(a_fecha))
    d["_mes"] = d["_ffact"].dt.to_period("M")
    # Mes en el que se muestra: lo que no está PAGADA ni CANCELADA se arrastra al mes actual
    # (solo en pantalla; la fecha de facturación no cambia)
    actual = pd.Period(date.today(), "M")
    arrastra = ~d["_est"].isin(["PAGADA", "CANCELADA"]) & (d["_mes"].isna() | (d["_mes"] < actual))
    d["_grupo"] = d["_mes"].where(~arrastra, actual)
    d["_arrastrada"] = arrastra
    d["_completo"] = df.apply(es_completo, axis=1)
    d["_con_xml"] = d["Número Factura"].fillna("").astype(str).str.strip() != ""
    return d


def ordenar(d: pd.DataFrame) -> pd.DataFrame:
    """Mes en que se muestra (reciente primero, sin fecha arriba) y estatus."""
    mes = d["_grupo"].map(lambda p: p.year * 100 + p.month if pd.notna(p) else 999999)
    est = d["_est"].map(ESTATUS_ORDEN).fillna(len(ESTATUS_ORDEN))
    return (d.assign(_o1=mes, _o2=est)
             .sort_values(["_o1", "_o2"], ascending=[False, True], kind="stable")
             .drop(columns=["_o1", "_o2"]))


def mostrar_flash():
    f = st.session_state.pop("flash", None)
    if f:
        getattr(st, f[0])(f[1])


def tras_guardar(tipo: str, msg: str):
    """Mensaje para la siguiente pantalla y reinicio de la selección de la tabla."""
    st.session_state["flash"] = (tipo, msg)
    st.session_state["ver"] = st.session_state.get("ver", 0) + 1
    st.rerun()


def tarjeta(col, titulo, valor, nota="", ayuda=None):
    with col.container(border=True):
        st.metric(titulo, valor, help=ayuda)
        if nota:
            st.caption(nota)


def excel(hojas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre_hoja, d in hojas.items():
            d.to_excel(w, index=False, sheet_name=nombre_hoja)
    return buf.getvalue()


XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ----------------------------------------------------------------------------
# USUARIOS Y LOGIN (mismos usuarios que la V1)
# ----------------------------------------------------------------------------
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
    cookie_name="facturacion_finarq_v2",
    key="clave_secreta_finarq_2026",
    cookie_expiry_days=1,
)


# ----------------------------------------------------------------------------
# PÁGINA: REGISTROS (una tabla por mes)
# ----------------------------------------------------------------------------
def _como_fecha(v):
    """date | Timestamp | texto | vacío -> date o None."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if hasattr(v, "date") and not isinstance(v, date):
        return v.date()
    if isinstance(v, date):
        return v
    return a_fecha(v)


def _fechas(s: pd.Series) -> pd.Series:
    return s.map(lambda x: x.date() if pd.notna(x) else None).astype(object)


def tabla_lectura(d: pd.DataFrame, todas: bool) -> pd.DataFrame:
    """Tabla de solo consulta (perfil comercial)."""
    t = pd.DataFrame({
        "Estatus": d["_est"].map(etiqueta),
        "Cliente": d["CLIENTE"],
        "Concepto": d["CONCEPTO"],
        "Producto": d["PRODUCTO"],
        "Asesor": d["Asesor"],
        "Total": d["_total"],
        "Factura": d["Número Factura"].where(d["_con_xml"], "— sin XML"),
        "Facturación": d["_ffact"],
        "Pago programado": d["_pp"],
        "Fecha de pago": d["_fpago"],
        "Comprobante de pago": d["Comprobante de Pago"],
        "Datos": d["_completo"].map({True: "✅ Completo", False: "✏️ Incompleto"}),
    })
    if todas:
        extra = pd.DataFrame({
            "Código SAT": d["CÓDIGO SAT"], "Aliado": d["Aliado Comercial"],
            "Volumen": d["_base"], "Subtotal": d["_sub"], "IVA": d["_iva"],
            "Folio": d["FOLIO (LISTO P/FACTURA)"], "Método de pago": d["Método de Pago"],
            "Complemento de pago": d["Complemento de Pago"], "Notas": d["NOTAS"],
            "Capturó": d["Capturado_Por"],
        })
        t = pd.concat([t, extra], axis=1)
    t["ID"] = d["ID"]
    return t.reset_index(drop=True)


CFG_LECTURA = {
    "Total": st.column_config.NumberColumn(format="dollar"),
    "Volumen": st.column_config.NumberColumn(format="dollar"),
    "Subtotal": st.column_config.NumberColumn(format="dollar"),
    "IVA": st.column_config.NumberColumn(format="dollar"),
    "Facturación": st.column_config.DateColumn(format="DD/MM/YYYY"),
    "Pago programado": st.column_config.DateColumn(
        format="DD/MM/YYYY", help="PUE: fin del mes de emisión · PPD: 60 días después de la emisión"),
    "Fecha de pago": st.column_config.DateColumn(format="DD/MM/YYYY"),
    "Comprobante de pago": st.column_config.TextColumn(width="medium"),
    "Concepto": st.column_config.TextColumn(width="medium"),
    "Cliente": st.column_config.TextColumn(width="medium"),
    "ID": st.column_config.TextColumn(width="small"),
}


def _color_fila(col_estatus: str):
    """Pinta la fila completa según su estatus (como en la V1)."""
    def pintar(row):
        est = ETIQUETA_A_ESTADO.get(row[col_estatus], row[col_estatus])
        css = f"background-color: {ESTADOS[est]['tinte']}" if est in ESTADOS else ""
        return [css] * len(row)
    return pintar


# --- Tabla editable (Administración): misma lógica que la V1 -----------------
# Se editan las mismas columnas que en la V1, más Comprobante y Complemento de pago (la
# conciliación los sigue llenando sola). IVA y TOTAL no se editan: se recalculan
# solos si cambia el SUBTOTAL. VENCIDA no se guarda (se conserva el estatus anterior).
FECHA_COLS = ["FECHA DE FACTURACION", "FECHA DE PAGO"]
NUM_COLS = ["BASE (Volumen)", "SUBTOTAL (INGRESO)"]
EDIT_COLS = ["ESTATUS DE PAGO", "CLIENTE", "CONCEPTO", "PRODUCTO", "Asesor", "Aliado Comercial",
             "CÓDIGO SAT", "BASE (Volumen)", "SUBTOTAL (INGRESO)", "FOLIO (LISTO P/FACTURA)",
             "FECHA DE FACTURACION", "FECHA DE PAGO", "Comprobante de Pago", "Complemento de Pago", "NOTAS"]
EXTRA_EDIT = ["Aliado Comercial", "CÓDIGO SAT", "FOLIO (LISTO P/FACTURA)", "Complemento de Pago", "NOTAS"]


def tabla_edicion(d: pd.DataFrame, todas: bool) -> pd.DataFrame:
    t = pd.DataFrame({
        "ESTATUS DE PAGO": d["_est"],
        "CLIENTE": d["CLIENTE"], "CONCEPTO": d["CONCEPTO"],
        "PRODUCTO": d["PRODUCTO"], "Asesor": d["Asesor"],
        "BASE (Volumen)": d["_base"], "SUBTOTAL (INGRESO)": d["_sub"],
        "IVA": d["_iva"], "TOTAL": d["_total"],
        "FECHA DE FACTURACION": _fechas(d["_ffact"]),
        "FECHA DE PAGO": _fechas(d["_fpago"]),
        "Comprobante de Pago": d["Comprobante de Pago"],
        "PAGO PROGRAMADO": _fechas(d["_pp"]),
        "Factura": d["Número Factura"].where(d["_con_xml"], "— sin XML"),
    })
    if todas:
        for c in EXTRA_EDIT:
            t[c] = d[c]
        t["Método de Pago"] = d["Método de Pago"]
        t["Capturado_Por"] = d["Capturado_Por"]
    t["ID"] = d["ID"]
    t["🗑️ Eliminar"] = False
    return t.reset_index(drop=True)


def _cfg_edicion(d: pd.DataFrame) -> dict:
    def ops(base, col):
        return list(dict.fromkeys([x.upper() for x in base]
                                  + sorted(x for x in d[col].dropna().unique() if str(x).strip())))
    ro = lambda label, **kw: st.column_config.Column(label, disabled=True, **kw)
    return {
        "ESTATUS DE PAGO": st.column_config.SelectboxColumn("Estatus", options=ESTATUS_VISTA, required=True,
                                                            help="VENCIDA se calcula sola; no se guarda"),
        "CLIENTE": st.column_config.TextColumn("Cliente", width="medium"),
        "CONCEPTO": st.column_config.TextColumn("Concepto", width="medium"),
        "PRODUCTO": st.column_config.SelectboxColumn("Producto", options=ops(PRODUCTOS, "PRODUCTO")),
        "Asesor": st.column_config.SelectboxColumn("Asesor", options=ops(ASESORES, "Asesor")),
        "BASE (Volumen)": st.column_config.NumberColumn("Volumen", min_value=0.0, format="dollar"),
        "SUBTOTAL (INGRESO)": st.column_config.NumberColumn("Subtotal", min_value=0.0, format="dollar"),
        "IVA": st.column_config.NumberColumn("IVA", disabled=True, format="dollar"),
        "TOTAL": st.column_config.NumberColumn("Total", disabled=True, format="dollar"),
        "FECHA DE FACTURACION": st.column_config.DateColumn("Facturación", format="DD/MM/YYYY"),
        "FECHA DE PAGO": st.column_config.DateColumn("Fecha de pago", format="DD/MM/YYYY"),
        "PAGO PROGRAMADO": st.column_config.DateColumn(
            "Pago programado", disabled=True, format="DD/MM/YYYY",
            help="PUE: fin del mes de emisión · PPD: 60 días después de la emisión"),
        "Factura": ro("Factura"),
        "Aliado Comercial": st.column_config.TextColumn("Aliado"),
        "CÓDIGO SAT": st.column_config.TextColumn("Código SAT"),
        "FOLIO (LISTO P/FACTURA)": st.column_config.TextColumn("Folio"),
        "NOTAS": st.column_config.TextColumn("Notas"),
        "Comprobante de Pago": st.column_config.TextColumn(
            "Comprobante de pago", width="medium",
            help="Lo llena la conciliación; escríbelo a mano si el pago no vino del estado de cuenta"),
        "Complemento de Pago": st.column_config.TextColumn(
            "Complemento de pago", help="Folio del CFDI de complemento de pago (facturas PPD)"),
        "Método de Pago": ro("Método de pago"),
        "Capturado_Por": ro("Capturó"),
        "ID": ro("ID", width="small"),
        "🗑️ Eliminar": st.column_config.CheckboxColumn("🗑️ Eliminar"),
    }


def _normaliza(col, v):
    """Valor editado -> valor comparable; None si no se debe guardar."""
    if col == "ESTATUS DE PAGO" and v == "VENCIDA":
        return None
    if col in FECHA_COLS:
        return _como_fecha(v)
    if col in NUM_COLS:
        n = pd.to_numeric(v, errors="coerce")
        return None if pd.isna(n) else float(n)
    v = "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip()
    return v.upper() if col in UPPER_COLS else v


def _difiere(col, nuevo, viejo) -> bool:
    if col in FECHA_COLS:
        return nuevo != _como_fecha(viejo)
    if col in NUM_COLS:
        viejo = pd.to_numeric(viejo, errors="coerce")
        if nuevo is None or pd.isna(viejo):
            return not (nuevo is None and pd.isna(viejo))
        return abs(nuevo - float(viejo)) >= 0.005
    return not _igual(nuevo, viejo)


def editor_mes(sub: pd.DataFrame, ref: pd.DataFrame, cfg: dict, todas: bool, mes: str) -> pd.DataFrame:
    """Tabla editable de un mes. Los cambios se acumulan en session_state['pend'] por ID,
    así sobreviven a cambios de filtros o de página hasta que se guardan o descartan."""
    pend = st.session_state.setdefault("pend", {})
    ver = st.session_state.get("ver", 0)
    orig = tabla_edicion(sub, todas)
    firma = abs(hash((ver, todas, mes, int(pd.util.hash_pandas_object(orig.astype(str), index=False).sum()))))

    # La tabla que recibe el editor solo se rearma cuando cambian los filtros o los datos;
    # si cambiara en cada edición, Streamlit reiniciaría el editor y se perderían cambios.
    snaps = st.session_state.setdefault("snaps", {})
    snap = snaps.get(mes)
    if not snap or snap[0] != firma:
        if snap:  # rescatar lo editado en la tabla anterior que aún no se había procesado
            previo = st.session_state.get(f"ed_{snap[0]}") or {}
            for fila, cols in previo.get("edited_rows", {}).items():
                rid = snap[1].iloc[int(fila)]["ID"]
                for col, v in cols.items():
                    if col in EDIT_COLS and rid in ref.index:
                        v = _normaliza(col, v)
                        if v is not None and _difiere(col, v, ref.at[rid, col]):
                            pend.setdefault(rid, {})[col] = v
        base = orig.copy()
        for rid, cols in pend.items():
            idx = base.index[base["ID"] == rid]
            for col, v in cols.items():
                if col in base.columns:
                    base.loc[idx, col] = v
        snap = (firma, base)
        snaps[mes] = snap

    editado = st.data_editor(
        snap[1].style.apply(_color_fila("ESTATUS DE PAGO"), axis=1),
        hide_index=True, width="stretch", num_rows="fixed", column_config=cfg,
        height=min(38 + 35 * len(snap[1]), 460), key=f"ed_{firma}",
    )
    # Pendientes de estas filas = lo que difiere de lo guardado
    for _, r in editado.iterrows():
        rid = r["ID"]
        cambios = {}
        for col in EDIT_COLS:
            if col in editado.columns:
                v = _normaliza(col, r[col])
                if v is not None and _difiere(col, v, ref.at[rid, col]):
                    cambios[col] = v
            elif col in pend.get(rid, {}):
                cambios[col] = pend[rid][col]  # cambio en una columna que ahora está oculta
        if cambios:
            pend[rid] = cambios
        else:
            pend.pop(rid, None)
    return editado


def _limpiar_edicion():
    st.session_state["pend"] = {}
    st.session_state["snaps"] = {}
    st.session_state["ver"] = st.session_state.get("ver", 0) + 1


def guardar_pendientes():
    cambios = {}
    for rid, cols in st.session_state.get("pend", {}).items():
        c = {}
        for col, v in cols.items():
            if col in FECHA_COLS:
                c[col] = v.isoformat() if v else ""
            elif col in NUM_COLS:
                c[col] = "" if v is None else round(v, 2)
            else:
                c[col] = v
        # IVA y TOTAL solo se recalculan si cambió el SUBTOTAL de ese registro
        if "SUBTOTAL (INGRESO)" in c:
            sub = float(c["SUBTOTAL (INGRESO)"] or 0)
            c["IVA"] = round(sub * IVA_RATE, 2)
            c["TOTAL"] = round(sub + sub * IVA_RATE, 2)
        if "FECHA DE PAGO" in c:
            c["Pago Real"] = c["FECHA DE PAGO"]
        cambios[rid] = c
    faltan = actualizar_registros(cambios)
    _limpiar_edicion()
    if faltan:
        st.session_state["flash"] = ("warning", f"Cambios guardados, excepto en {', '.join(faltan)}: "
                                                f"ya no existen (alguien más los borró).")
    else:
        st.session_state["flash"] = ("success", f"✅ Cambios guardados en {len(cambios)} registro(s).")
    st.rerun()


def barra_cambios(borrar: list, donde: str):
    """Aviso de cambios sin guardar + botones (se muestra arriba y abajo de las tablas)."""
    pend = st.session_state.get("pend", {})
    if pend:
        ocultos = len(set(pend) - set(st.session_state.get("ids_visibles", [])))
        st.warning(f"✏️ Tienes cambios sin guardar en {len(pend)} registro(s)"
                   + (f" ({ocultos} no se ven con los filtros actuales)" if ocultos else "")
                   + ". Se conservan aunque cambies filtros o de página, pero se pierden si cierras "
                     "o recargas el navegador.")
        g, x, _ = st.columns([1, 1, 3])
        if g.button("💾 Guardar cambios", type="primary", key=f"guardar_{donde}", width="stretch"):
            guardar_pendientes()
        if x.button("↩️ Descartar cambios", key=f"descartar_{donde}", width="stretch"):
            _limpiar_edicion()
            st.rerun()
    if borrar:
        st.error(f"🗑️ Marcaste {len(borrar)} registro(s) para eliminar: {', '.join(borrar)}")
        c, b, _ = st.columns([1.3, 1, 2.7])
        ok = c.checkbox("Sí, eliminarlos", key=f"conf_borrar_{donde}")
        if b.button("Eliminar", disabled=not ok, key=f"borrar_{donde}", width="stretch"):
            eliminar_registros(borrar)
            for rid in borrar:
                st.session_state.get("pend", {}).pop(rid, None)
            st.session_state["snaps"] = {}
            tras_guardar("success", f"🗑️ {len(borrar)} registro(s) eliminado(s).")


def pagina_registros():
    st.title("📋 Registros")
    mostrar_flash()
    df = leer_registros()
    d = preparar(df)
    if d.empty:
        st.info("Aún no hay registros. Captúralos en «➕ Nuevo registro».")
        return

    def _opts(col):
        return sorted(x for x in d[col].dropna().unique() if str(x).strip())

    # Mismos filtros que la V1
    st.markdown("##### 🔎 Filtros")
    r1 = st.columns(4)
    f_est = r1[0].multiselect("Estatus de pago", ESTATUS_VISTA, format_func=etiqueta, key="reg_est")
    f_cli = r1[1].multiselect("Cliente", _opts("CLIENTE"), key="reg_cli")
    f_ase = r1[2].multiselect("Asesor", _opts("Asesor"), key="reg_ase")
    f_pro = r1[3].multiselect("Producto", _opts("PRODUCTO"), key="reg_pro")
    r2 = st.columns(3)
    f_ali = r2[0].multiselect("Aliado comercial", _opts("Aliado Comercial"), key="reg_ali")
    f_ffa = r2[1].date_input("Fecha de facturación (rango)", value=(), format="DD/MM/YYYY", key="reg_ffa")
    f_fpa = r2[2].date_input("Fecha de pago (rango)", value=(), format="DD/MM/YYYY", key="reg_fpa")

    m = pd.Series(True, index=d.index)
    if f_est: m &= d["_est"].isin(f_est)
    if f_cli: m &= d["CLIENTE"].isin(f_cli)
    if f_ase: m &= d["Asesor"].isin(f_ase)
    if f_pro: m &= d["PRODUCTO"].isin(f_pro)
    if f_ali: m &= d["Aliado Comercial"].isin(f_ali)
    if len(f_ffa) == 2: m &= d["_ffact"].between(pd.Timestamp(f_ffa[0]), pd.Timestamp(f_ffa[1]))
    if len(f_fpa) == 2: m &= d["_fpago"].between(pd.Timestamp(f_fpa[0]), pd.Timestamp(f_fpa[1]))
    hay_filtros = bool(f_est or f_cli or f_ase or f_pro or f_ali or len(f_ffa) == 2 or len(f_fpa) == 2)
    filt = ordenar(d[m])
    if filt.empty:
        st.info("No hay registros con esos filtros.")
        return

    t1, t2 = st.columns(2)
    todas = t1.toggle("Ver todas las columnas", key="reg_todas")
    sumar = not ES_ADMIN or t2.toggle("➕ Sumar como en Excel", key="reg_sumar",
                                      help="Las tablas dejan de ser editables y puedes seleccionar "
                                           "celdas para ver su suma. Apágalo para volver a editar.")
    editar = ES_ADMIN and not sumar
    if editar:
        st.caption("Edita directo en la tabla y presiona «💾 Guardar cambios». Marca 🗑️ para eliminar. "
                   "El color de la fila se actualiza al guardar.")
    else:
        st.caption("Selecciona importes como en Excel: arrastra, o usa Shift / Ctrl + clic. "
                   "La suma aparece arriba y abajo de las tablas.")
    arriba = st.container()
    if ES_ADMIN:
        ref = tabla_edicion(d, True).set_index("ID")
        cfg = _cfg_edicion(d)

    meses = filt["_grupo"].map(lambda p: nombre_mes(p) if pd.notna(p) else "Sin fecha de facturación")
    grupos = list(dict.fromkeys(meses))
    ver_todos = st.session_state.get("reg_todos_meses", False) or hay_filtros
    visibles = grupos if ver_todos else grupos[:MESES_VISIBLES]
    borrar, ids_visibles, valores = [], [], []
    for i, g in enumerate(visibles):
        sub = filt[meses == g]
        ids_visibles += sub["ID"].tolist()
        act = sub[sub["_est"] != "CANCELADA"]
        n = sub["_est"].value_counts()
        chips = "  ".join(f"{ESTADOS[e]['icono']} {n[e]}" for e in ESTATUS_VISTA if n.get(e))
        with st.expander(f"**{g}** — {len(sub)} registro(s) · {fmt_money(act['_total'].sum())} · {chips}",
                         expanded=(i == 0 or hay_filtros)):
            n_arr = int(sub["_arrastrada"].sum())
            if n_arr:
                st.caption(f"↪️ Incluye {n_arr} registro(s) de meses anteriores que aún no están pagados "
                           f"(revisa su fecha de facturación).")
            if editar:
                ed = editor_mes(sub, ref, cfg, todas, g)
                borrar += ed.loc[ed["🗑️ Eliminar"] == True, "ID"].tolist()
            else:
                valores += tabla_consulta(sub, todas, g)
    st.session_state["ids_visibles"] = ids_visibles

    ocultos = len(grupos) - len(visibles)
    if ocultos > 0:
        if st.button(f"Ver {ocultos} mes(es) anteriores"):
            st.session_state["reg_todos_meses"] = True
            st.rerun()

    with arriba:
        barra_suma(valores)
        if ES_ADMIN:
            barra_cambios(borrar, "arriba")
    barra_suma(valores)
    if ES_ADMIN:
        barra_cambios(borrar, "abajo")

    st.divider()
    c1, c2, _ = st.columns([1, 1, 2])
    c1.download_button("⬇️ Excel de esta vista", excel({"Registros": pd.concat(
        [filt["_est"].rename("ESTATUS DE PAGO"),
         filt[[c for c in LAYOUT_COLS if c != "ESTATUS DE PAGO"]],
         filt[["Comprobante de Pago", "Complemento de Pago"]],
         filt["_pp"].dt.strftime("%Y-%m-%d").fillna("").rename("PAGO PROGRAMADO")], axis=1)}),
        file_name="registros.xlsx", mime=XLSX, width="stretch")
    c2.download_button("⬇️ Registro Máster", excel({"Registro Maestro": layout_a_master(filt)}),
                       file_name="registro_master.xlsx", mime=XLSX, width="stretch")


def tabla_consulta(sub: pd.DataFrame, todas: bool, mes: str) -> list:
    """Tabla de solo consulta. Se pueden seleccionar celdas (clic, Shift o Ctrl) y se suman
    los importes, como en Excel. Regresa los valores seleccionados."""
    tabla = tabla_lectura(sub, todas)
    ev = st.dataframe(tabla.style.apply(_color_fila("Estatus"), axis=1),
                      hide_index=True, width="stretch", height=min(38 + 35 * len(tabla), 460),
                      column_config=CFG_LECTURA, on_select="rerun", selection_mode="multi-cell",
                      key=f"sel_{mes}")
    valores = []
    for fila, col in ev.selection.cells:
        v = tabla.iloc[fila][col] if col in tabla.columns else None
        if isinstance(v, (int, float)) and not isinstance(v, bool) and pd.notna(v):
            valores.append(float(v))
    return valores


def barra_suma(valores: list):
    """Suma de las celdas seleccionadas, como la barra de estado de Excel."""
    if valores:
        st.info(f"➕ **{len(valores)} celda(s)** · Suma **{fmt_money(sum(valores))}** · "
                f"Promedio {fmt_money(sum(valores) / len(valores))}")


# ----------------------------------------------------------------------------
# LIGAR FACTURA XML A UN REGISTRO (misma lógica que la V1)
# ----------------------------------------------------------------------------
def ligar_xml(df: pd.DataFrame, rid: str):
    reg = df.loc[df["ID"] == rid].iloc[0]
    if str(reg["Número Factura"]).strip():
        st.caption(f"Este registro ya tiene la factura **{reg['Número Factura']}**. "
                   "Si subes otra, se reemplazan los datos de factura.")
    xml_file = st.file_uploader("Sube el XML de la factura", type=["xml"], key=f"xml_{rid}")
    if xml_file is None:
        return
    try:
        datos = parse_cfdi(xml_file.getvalue())
    except Exception as ex:
        st.error(f"No se pudo leer el XML: {ex}")
        return
    cliente_reg = str(reg["CLIENTE"])
    datos["Concepto (Factura)"] = match_concepto(cliente_reg, datos["_Conceptos"])

    c1, c2 = st.columns(2)
    c1.markdown("**Lo que dice la factura**")
    c1.markdown(f"Número: **{datos['Número Factura']}**  \nEmisión: {datos['Fecha Emisión']}  \n"
                f"Método de pago: {datos['Método de Pago']}  \nFacturado a: {datos['Datos (Factura)']}  \n"
                f"Concepto: {datos['Concepto (Factura)']}  \nTotal: **{fmt_money(datos['_Total'])}**")
    c2.markdown("**Lo que dice el registro**")
    c2.markdown(f"Cliente: {cliente_reg}  \nTotal: **{fmt_money(reg['TOTAL'])}**")
    if len(datos["_Conceptos"]) > 1:
        c2.caption(f"La factura trae {len(datos['_Conceptos'])} descripciones; se tomó la que coincide con «{cliente_reg}».")
    try:
        if datos["_Total"] and abs(float(datos["_Total"]) - float(reg["TOTAL"])) > 0.5:
            st.warning("⚠️ El total de la factura no coincide con el del registro. "
                       "Al guardar, Subtotal, IVA y Total se corrigen con los de la factura.")
    except Exception:
        pass
    st.caption(f"UUID: {datos['_UUID']}")
    if st.button("💾 Ligar factura al registro", type="primary", key=f"ligar_{rid}"):
        cambios, corregido = aplicar_factura(reg, datos)
        actualizar_registros({rid: cambios})
        if corregido:
            tras_guardar("warning", f"Factura {datos['Número Factura']} ligada a {rid}. El total no coincidía: "
                                    f"Subtotal, IVA y Total se corrigieron con el XML (Total {fmt_money(datos['_Total'])}).")
        tras_guardar("success", f"✅ Factura {datos['Número Factura']} ligada a {rid}.")


# ----------------------------------------------------------------------------
# PÁGINA: NUEVO REGISTRO (mismos formularios y cálculos que la V1)
# ----------------------------------------------------------------------------
def pagina_nuevo():
    st.title("➕ Nuevo registro")
    mostrar_flash()
    df = leer_registros()

    if ES_ADMIN:
        st.caption("Captura completa. IVA y Total se calculan solos.")
        with st.form("form_admin", clear_on_submit=True):
            st.markdown("**Cliente y producto**")
            c1, c2, c3 = st.columns(3)
            cliente = c1.text_input("Cliente *")
            producto = c2.selectbox("Producto *", PRODUCTOS)
            asesor = c3.selectbox("Asesor *", ASESORES)
            c1, c2, c3 = st.columns(3)
            concepto = c1.text_input("Concepto")
            aliado = c2.text_input("Aliado comercial")
            codigo_sat = c3.text_input("Código SAT")
            st.markdown("**Montos**")
            c1, c2, c3 = st.columns(3)
            base = c1.number_input("Volumen (base)", min_value=0.0, step=100.0, format="%.2f")
            tasa = c2.number_input("Tasa (%) — opcional", min_value=0.0, max_value=100.0, step=0.5, format="%.2f",
                                   help="Si la llenas, Subtotal = Volumen × Tasa / 100")
            subtotal_in = c3.number_input("Subtotal (ingreso)", min_value=0.0, step=100.0, format="%.2f",
                                          help="Se usa solo si la tasa es 0")
            st.markdown("**Estatus**")
            c1, c2 = st.columns([1, 2])
            estatus = c1.selectbox("Estatus", ESTATUS, format_func=etiqueta)
            notas = c2.text_input("Notas")
            enviar = st.form_submit_button("💾 Guardar registro", type="primary", width="stretch")

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
                    "ESTATUS DE PAGO": estatus, "NOTAS": notas.strip(),
                    "ID": siguiente_id(df), "Capturado_Por": nombre,
                })
                guardar_registro(fila)
                st.session_state["flash"] = ("success", f"✅ Registro {fila['ID']} guardado · "
                                                        f"{cliente.strip().upper()} · Total {fmt_money(fila['TOTAL'])}")
                st.rerun()
    else:
        st.caption("Captura tus movimientos. Administración completará el resto de la información.")
        producto = st.selectbox("Producto", PRODUCTOS, key="prod_com")
        es_fx = producto == "FX"
        clientes_fx = cargar_clientes_fx()
        if es_fx:
            st.caption(f"Para FX se llenan solos el aliado ({FX_ALIADO}) y el código SAT ({FX_CODIGO_SAT}).")

        with st.form("form_comercial", clear_on_submit=True):
            c1, c2 = st.columns(2)
            if es_fx:
                cliente_sel = c1.selectbox("Cliente *", clientes_fx + ["➕ Otro (agregar)…"])
                cliente_nuevo = c2.text_input("Nuevo cliente (solo si elegiste «Otro»)")
            else:
                cliente_sel = c1.text_input("Cliente *")
                cliente_nuevo = ""
            c1, c2 = st.columns(2)
            asesor = c1.selectbox("Asesor *", ASESORES)
            aliado = c2.text_input("Aliado comercial", value=FX_ALIADO if es_fx else "")
            concepto = st.text_input("Concepto")
            c1, c2, c3 = st.columns(3)
            base = c1.number_input("Volumen", min_value=0.0, step=100.0, format="%.2f")
            tasa = c2.number_input("Tasa (%)", min_value=0.0, max_value=100.0, step=0.5, format="%.2f",
                                   help="Subtotal (Ingreso) = Volumen × Tasa / 100")
            codigo_sat = c3.text_input("Código SAT", value=FX_CODIGO_SAT if es_fx else "")
            enviar = st.form_submit_button("💾 Guardar registro", type="primary", width="stretch")

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
                st.session_state["flash"] = ("success", f"✅ Registro {fila['ID']} guardado. Administración "
                                                        f"completará los datos que falten.")
                st.rerun()

    st.page_link(PAGINAS["registros"], label="Ver registros", icon="📋")


# ----------------------------------------------------------------------------
# PÁGINA: FACTURAS (XML)
# ----------------------------------------------------------------------------
def pagina_facturas():
    st.title("📎 Facturas (XML)")
    mostrar_flash()
    df = leer_registros()
    d = preparar(df)

    st.subheader("Ligar una factura a un registro")
    if d.empty:
        st.info("Primero captura movimientos.")
    else:
        incluir = st.toggle("Incluir registros que ya tienen factura", key="fx_incluir")
        cand = ordenar(d if incluir else d[~d["_con_xml"]])
        if cand.empty:
            st.success("Todos los registros ya tienen su factura ligada. 🎉")
        else:
            etiquetas = {r["ID"]: f"{r['ID']} · {r['CLIENTE']} · {fmt_money(r['_total'])} · {etiqueta(r['_est'])}"
                         for _, r in cand.iterrows()}
            rid = st.selectbox(f"Registro ({len(cand)} disponibles) — escribe para buscar",
                               list(etiquetas), format_func=etiquetas.get, index=None,
                               placeholder="Busca por cliente, ID o monto…", key="fx_rid")
            if rid:
                ligar_xml(df, rid)

    st.divider()
    st.subheader("Carga masiva de facturas Kushki")
    st.caption(f"Sube un ZIP o varios XML. Cada factura crea un registro nuevo con Producto «{KUSHKI_PRODUCTO}», "
               f"Aliado «{KUSHKI_ALIADO}» y Asesor «{KUSHKI_ASESOR}». Solo el Volumen se completa a mano.")
    lote = st.file_uploader("ZIP o XMLs", type=["zip", "xml"], accept_multiple_files=True,
                            key=f"lote_xml_{st.session_state.get('lote_n', 0)}")
    if not lote:
        return
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
        return
    prev_lote = pd.DataFrame([{
        "Crear": True, "Factura": dd["Número Factura"], "Cliente": dd["_Nombre"],
        "Concepto": dd["Concepto (Factura)"], "Emisión": dd["Fecha Emisión"],
        "Método": dd["Método de Pago"], "Total": pd.to_numeric(dd["_Total"], errors="coerce"),
        "Volumen": float("nan"), "Archivo": nom,
    } for nom, dd in nuevas])
    ro = ["Factura", "Cliente", "Concepto", "Emisión", "Método", "Archivo"]
    cfg = {c: st.column_config.TextColumn(c, disabled=True) for c in ro}
    cfg["Total"] = st.column_config.NumberColumn(format="dollar", disabled=True)
    cfg["Volumen"] = st.column_config.NumberColumn("Volumen ✏️", min_value=0.0, format="$%.2f",
                                                   help="Puedes dejarlo vacío y completarlo después")
    st.markdown(f"**{len(nuevas)} factura(s) nuevas.** Captura el Volumen (opcional) y crea los registros:")
    editado = st.data_editor(prev_lote, width="stretch", hide_index=True, num_rows="fixed",
                             key=f"editor_lote_{abs(hash(tuple(n for n, _ in nuevas)))}", column_config=cfg)
    if st.button("➕ Crear registros", type="primary"):
        full = leer_registros()
        n_id = int(siguiente_id(full).replace("MOV-", ""))
        filas = []
        for i, r in editado.iterrows():
            if r["Crear"] != True:
                continue
            fila = fila_desde_xml(nuevas[i][1], f"MOV-{n_id:05d}", nombre)
            n_id += 1
            if pd.notna(r["Volumen"]):
                fila["BASE (Volumen)"] = r["Volumen"]
            filas.append(fila)
        if filas:
            agregar_registros(filas)
            st.session_state["lote_n"] = st.session_state.get("lote_n", 0) + 1
            tras_guardar("success", f"✅ {len(filas)} registro(s) creados desde XML "
                                    f"({filas[0]['ID']} a {filas[-1]['ID']}).")
        else:
            st.info("No marcaste ninguna factura.")


# ----------------------------------------------------------------------------
# PÁGINA: CONCILIAR PAGOS (estado de cuenta; misma lógica que la V1)
# ----------------------------------------------------------------------------
def pagina_conciliar():
    st.title("🏦 Conciliar pagos")
    st.caption("Sube el PDF del estado de cuenta. Cada depósito se busca entre los registros por su monto "
               "total; tú eliges cuáles marcar como pagados.")
    mostrar_flash()
    df = leer_registros()
    if df.empty:
        st.info("Aún no hay registros.")
        return
    edc = st.file_uploader("PDF del estado de cuenta", type=["pdf"])
    if edc is None:
        return
    try:
        deps = parse_estado_cuenta(edc.getvalue())
    except Exception as ex:
        st.error(f"No se pudo leer el estado de cuenta: {ex}")
        deps = []

    filas = []
    for _, r in df[df["ESTATUS DE PAGO"] != "CANCELADA"].iterrows():  # las canceladas no se concilian
        dep = match_deposito(r["TOTAL"], deps)
        if dep:
            fp_reg = _fecha_iso(r["FECHA DE PAGO"])
            filas.append({
                "Aplicar": True, "ID": r["ID"], "Cliente": r["CLIENTE"],
                "Total": pd.to_numeric(r["TOTAL"], errors="coerce"),
                "Estatus actual": etiqueta(str(r["ESTATUS DE PAGO"]).upper()),
                "Fecha en registro": fp_reg or "—",
                "Fecha del depósito": dep["fecha"],
                "Revisión": ("✅ Coincide" if fp_reg == dep["fecha"] else
                             "📝 Se registrará la fecha" if not fp_reg else "⚠️ Se corregirá la fecha"),
                "Comprobante": dep["concepto"],
                "_pago_real": dep["fecha"], "_comprobante": dep["concepto"], "_fp_reg": fp_reg,
            })
    totales = pd.to_numeric(df.loc[df["ESTATUS DE PAGO"] != "CANCELADA", "TOTAL"], errors="coerce").dropna().tolist()
    n_dep_ok = sum(1 for x in deps if any(abs(x["monto"] - t) < 0.01 for t in totales))
    k = st.columns(3)
    tarjeta(k[0], "Depósitos en el PDF", str(len(deps)), fmt_money(sum(x["monto"] for x in deps)))
    tarjeta(k[1], "Coinciden con un registro", str(len(filas)))
    tarjeta(k[2], "Sin coincidencia", str(len(deps) - n_dep_ok),
            "Depósitos cuyo monto no está en ningún registro")
    if not filas:
        st.warning("Ningún registro coincidió por monto con los depósitos del estado de cuenta.")
        return

    prev = pd.DataFrame(filas)
    n_dif = int((prev["Revisión"] == "⚠️ Se corregirá la fecha").sum())
    if n_dif:
        st.warning(f"⚠️ {n_dif} registro(s) tienen una fecha de pago distinta a la del depósito. "
                   "Al aplicar se usa la fecha del depósito.")
    vis = ["ID", "Cliente", "Total", "Estatus actual", "Fecha en registro", "Fecha del depósito",
           "Revisión", "Comprobante"]
    cfg = {c: st.column_config.TextColumn(c, disabled=True) for c in vis}
    cfg["Total"] = st.column_config.NumberColumn(format="dollar", disabled=True)
    cfg["Aplicar"] = st.column_config.CheckboxColumn("Aplicar", help="Desmarca los que no quieras aplicar")
    editado = st.data_editor(prev[["Aplicar"] + vis], width="stretch", hide_index=True,
                             num_rows="fixed", key="editor_edc", column_config=cfg)
    aplicar = editado.loc[editado["Aplicar"] == True, "ID"].tolist()
    if st.button(f"💾 Marcar {len(aplicar)} como pagados", type="primary", disabled=not aplicar):
        ref = prev.set_index("ID")
        cambios, corregidos = {}, []
        for rid in aplicar:
            pago_real = ref.loc[rid, "_pago_real"]
            cambios[rid] = {"Pago Real": pago_real, "Comprobante de Pago": ref.loc[rid, "_comprobante"],
                            "ESTATUS DE PAGO": "PAGADA"}
            if ref.loc[rid, "_fp_reg"] != pago_real:
                cambios[rid]["FECHA DE PAGO"] = pago_real
                corregidos.append(rid)
        actualizar_registros(cambios)
        msg = f"✅ Se aplicó el pago a {len(aplicar)} registro(s)."
        if corregidos:
            msg += f" Se corrigió la fecha de pago en: {', '.join(corregidos)}."
        tras_guardar("warning" if corregidos else "success", msg)


# ----------------------------------------------------------------------------
# ARRANQUE: login, menú y página
# ----------------------------------------------------------------------------
authenticator.login("Iniciar sesión", location="main")

if not st.session_state.get("authentication_status"):
    st.title("🧾 Facturación FIN.ARQ")
    if st.session_state.get("authentication_status") is False:
        st.error("Usuario o contraseña incorrectos")
    else:
        st.info("Ingresa tus credenciales para continuar.")
    st.stop()

usuario = st.session_state["username"]
rol = credentials["usernames"][usuario]["role"]
nombre = st.session_state["name"]
ES_ADMIN = rol == "admin"

PAGINAS = {
    "registros": st.Page(pagina_registros, title="Registros", icon="📋", default=ES_ADMIN),
    "nuevo": st.Page(pagina_nuevo, title="Nuevo registro", icon="➕", default=not ES_ADMIN),
    "facturas": st.Page(pagina_facturas, title="Facturas (XML)", icon="📎"),
    "conciliar": st.Page(pagina_conciliar, title="Conciliar pagos", icon="🏦"),
}
visibles = (["registros", "nuevo", "facturas", "conciliar"] if ES_ADMIN
            else ["nuevo", "registros"])

with st.sidebar:
    st.markdown("## 🧾 Facturación FIN.ARQ")
    st.caption("Versión 2 · demo")
nav = st.navigation([PAGINAS[k] for k in visibles])
with st.sidebar:
    st.divider()
    st.markdown(f"👤 **{nombre}**")
    if st.button("🔄 Actualizar datos", width="stretch",
                 help="Vuelve a leer la hoja (útil si alguien más acaba de capturar)"):
        leer_registros.clear()
        st.rerun()
    if ES_ADMIN and _usar_sheets():
        st.link_button("📗 Abrir Google Sheets", url_sheet(), width="stretch")
    authenticator.logout("Cerrar sesión", location="sidebar")
    st.caption("Conectado a Google Sheets ✅" if _usar_sheets() else "⚠️ Modo local (CSV de prueba)")
    with st.expander("¿Qué significa cada estatus?"):
        for e in ESTATUS_VISTA:
            st.markdown(f"**{etiqueta(e)}**  \n{ESTADOS[e]['ayuda']}")
        st.caption("VENCIDA no se guarda en la hoja: se calcula sola.")

aviso = st.sidebar.empty()
nav.run()
if st.session_state.get("pend"):
    aviso.warning(f"✏️ {len(st.session_state['pend'])} registro(s) con cambios sin guardar en 📋 Registros")
