import streamlit as st
import io
import os
import sys
import tempfile

# ── Importar el motor de generación ─────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from generar_cuadros_cocodi import (
    detectar_fecha, textos_fecha, leer_csv_map, leer_xlsx_map,
    hoja_pp, hoja_pp_cap, hoja_agricultura, hoja_comisario,
    MESES_ES, DIAS_FIN_MES,
    TPL_PP, TPL_PPCAP, TPL_AGR, TPL_COM,
)
import openpyxl

# ── Configuración de página ──────────────────────────────────────────
st.set_page_config(
    page_title="Cuadros COCODI — SADER",
    page_icon="📊",
    layout="centered",
)

# ── Estilo ───────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { max-width: 720px; padding-top: 2rem; }
    .stDownloadButton > button {
        background-color: #9F2241;
        color: white;
        font-weight: 600;
        border: none;
        width: 100%;
        padding: 0.6rem 1.2rem;
        border-radius: 6px;
        font-size: 1rem;
    }
    .stDownloadButton > button:hover { background-color: #7a1a32; }
    .stButton > button {
        background-color: #BC955C;
        color: white;
        font-weight: 600;
        border: none;
        width: 100%;
        padding: 0.6rem 1.2rem;
        border-radius: 6px;
        font-size: 1rem;
    }
    .stButton > button:hover { background-color: #9a7840; }
    h1 { color: #9F2241; }
    .info-box {
        background-color: #f9f3eb;
        border-left: 4px solid #BC955C;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        margin-bottom: 1rem;
        font-size: 0.9rem;
        color: #444;
    }
</style>
""", unsafe_allow_html=True)

# ── Encabezado ───────────────────────────────────────────────────────
st.title("Generador de Cuadros COCODI")
st.markdown(
    "<div class='info-box'>"
    "Sube el archivo MAP del día (<b>.csv</b> o <b>.xlsx</b> con hojas TD) "
    "y descarga automáticamente el Excel con las 4 hojas: "
    "<b>Pp</b>, <b>Pp y CAP</b>, <b>AGRICULTURA</b> y <b>Presupuesto Comisario</b>."
    "</div>",
    unsafe_allow_html=True,
)

# ── Upload ───────────────────────────────────────────────────────────
archivo = st.file_uploader(
    "Archivo MAP del día",
    type=["csv", "xlsx"],
    help="Nombre sugerido: DD-MES-AAAA_MAP.csv  —  el script detecta la fecha automáticamente.",
)

if archivo is not None:
    nombre = archivo.name
    ext = os.path.splitext(nombre)[1].lower()

    if st.button("Generar cuadros"):
        with st.spinner("Procesando…"):
            try:
                # Guardar el archivo subido en un temp
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext,
                                                  prefix="MAP_") as tmp:
                    tmp.write(archivo.getvalue())
                    ruta_tmp = tmp.name

                # Detectar fecha del nombre original
                fecha = detectar_fecha(nombre)
                tf    = textos_fecha(fecha)

                # Leer datos
                if ext == ".csv":
                    datos = leer_csv_map(ruta_tmp, fecha["mes"])
                else:
                    datos = leer_xlsx_map(ruta_tmp)
                datos["_mes"] = fecha["mes"]

                # Generar workbook
                tpls = {"Pp": TPL_PP, "PpCAP": TPL_PPCAP,
                        "AGRICULTURA": TPL_AGR, "Comisario": TPL_COM}

                wb_out = openpyxl.Workbook()
                if "Sheet" in wb_out.sheetnames:
                    del wb_out["Sheet"]

                hoja_pp(wb_out, datos, tf, tpls["Pp"])
                hoja_pp_cap(wb_out, datos, tf, tpls["PpCAP"])
                hoja_agricultura(wb_out, datos, tf, tpls["AGRICULTURA"])
                hoja_comisario(wb_out, datos, tf, tpls["Comisario"])

                # Serializar a bytes
                buf = io.BytesIO()
                wb_out.save(buf)
                buf.seek(0)

                mes_str    = MESES_ES[fecha["mes"]].upper()
                nombre_out = (f"CUADROS_COCODI_"
                              f"{fecha['dia']:02d}-{mes_str}-{fecha['anio']}.xlsx")

                os.unlink(ruta_tmp)

                st.success(f" Cuadros generados — corte al {tf['dia_mes_anio']}")

                st.download_button(
                    label=f"⬇️  Descargar {nombre_out}",
                    data=buf,
                    file_name=nombre_out,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
                st.exception(e)
