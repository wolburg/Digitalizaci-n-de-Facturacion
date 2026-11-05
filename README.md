# Digitalización de Facturación — FIN.ARQ

Herramienta web (Streamlit) para el registro y seguimiento de facturación:
captura de movimientos (LayOut), conciliación con facturas (XML CFDI) y con el
estado de cuenta bancario (PDF).

## Perfiles
- **comercial**: captura movimientos (con lógica FX) y consulta registros.
- **administracion**: LayOut completo, facturación (XML) y estado de cuenta.

## Instalación
```bash
pip install -r requirements.txt
streamlit run LayOut.py
```
Sin credenciales de Google, la app corre en **modo local (CSV)** y guarda en `data/`.

## Estructura
- `LayOut.py` — aplicación principal (versión 2): una tabla por mes, edición directa y estatus CANCELADA.
- `core.py` — reglas y capa de datos que usa `LayOut.py`.
- `LayOut_v1.py` — versión 1, solo como respaldo. Para regresar a ella, intercambiar los nombres con `LayOut.py`.
- `requirements.txt` — dependencias.
- `data/` — datos locales (no se sube al repo).

## Producción (Google Sheets)
Configurar una cuenta de servicio de Google y poner las credenciales en
`.streamlit/secrets.toml` (`sheet_url` + `[gcp_service_account]`). Este archivo
NO debe subirse al repositorio.
