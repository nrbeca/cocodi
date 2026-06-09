import streamlit as st
import io
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(__file__))

from generar_cuadros_cocodi import (
    detectar_fecha, textos_fecha, leer_csv_map, leer_xlsx_map,
    hoja_pp, hoja_pp_cap, hoja_agricultura, hoja_comisario,
    MESES_ES, DIAS_FIN_MES,
    TPL_PP, TPL_PPCAP, TPL_AGR, TPL_COM,
)
from generar_austeridad import (
    leer_cp2025, leer_sicop2026, detectar_corte, generar_austeridad,
)
import openpyxl

# ── Configuración ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Reportes COCODI — SADER",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
    .block-container { max-width: 740px; padding-top: 2rem; }

    /* Botón descargar */
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

    /* Botón generar */
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
    h2 { color: #9F2241; margin-top: 0; }

    .info-box {
        background-color: #f9f3eb;
        border-left: 4px solid #BC955C;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        margin-bottom: 1rem;
        font-size: 0.9rem;
        color: #444;
    }
    .section-divider {
        border: none;
        border-top: 2px solid #e8d5b0;
        margin: 2.5rem 0 2rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════
#  SECCIÓN 1 — CUADROS COCODI
# ════════════════════════════════════════════════════════════════════
st.title("Reportes COCODI — SADER")

st.markdown("##  Generador de Cuadros COCODI")
st.markdown(
    "<div class='info-box'>"
    "Sube el archivo MAP del día (<b>.csv</b> o <b>.xlsx</b>) "
    "y descarga el Excel con las 4 hojas: "
    "<b>Pp</b>, <b>Pp y CAP</b>, <b>AGRICULTURA</b> y <b>Presupuesto Comisario</b>."
    "</div>",
    unsafe_allow_html=True,
)

archivo_map = st.file_uploader(
    "Archivo MAP del día",
    type=["csv", "xlsx"],
    key="uploader_cocodi",
    help="Nombre sugerido: DD-MES-AAAA_MAP.csv — la fecha se detecta automáticamente.",
)

if archivo_map is not None:
    nombre_map = archivo_map.name
    ext_map    = os.path.splitext(nombre_map)[1].lower()

    if st.button("Generar cuadros COCODI", key="btn_cocodi"):
        with st.spinner("Procesando…"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext_map, prefix="MAP_") as tmp:
                    tmp.write(archivo_map.getvalue())
                    ruta_tmp = tmp.name

                fecha = detectar_fecha(nombre_map)
                tf    = textos_fecha(fecha)

                if ext_map == ".csv":
                    datos = leer_csv_map(ruta_tmp, fecha["mes"])
                else:
                    datos = leer_xlsx_map(ruta_tmp)
                datos["_mes"] = fecha["mes"]

                tpls   = {"Pp": TPL_PP, "PpCAP": TPL_PPCAP,
                          "AGRICULTURA": TPL_AGR, "Comisario": TPL_COM}
                wb_out = openpyxl.Workbook()
                if "Sheet" in wb_out.sheetnames:
                    del wb_out["Sheet"]

                hoja_pp(wb_out, datos, tf, tpls["Pp"])
                hoja_pp_cap(wb_out, datos, tf, tpls["PpCAP"])
                hoja_agricultura(wb_out, datos, tf, tpls["AGRICULTURA"])
                hoja_comisario(wb_out, datos, tf, tpls["Comisario"])

                buf = io.BytesIO()
                wb_out.save(buf)
                buf.seek(0)

                mes_str    = MESES_ES[fecha["mes"]].upper()
                nombre_out = f"CUADROS_COCODI_{fecha['dia']:02d}-{mes_str}-{fecha['anio']}.xlsx"

                os.unlink(ruta_tmp)

                st.success(f" Cuadros generados — corte al {tf['dia_mes_anio']}")
                st.download_button(
                    label=f"⬇️  Descargar {nombre_out}",
                    data=buf,
                    file_name=nombre_out,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_cocodi",
                )

            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
                st.exception(e)

# ════════════════════════════════════════════════════════════════════
#  SECCIÓN 2 — FORMATO DE AUSTERIDAD
# ════════════════════════════════════════════════════════════════════
st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

st.markdown("##  Formato de Austeridad")
st.markdown(
    "<div class='info-box'>"
    "Sube la <b>Cuenta Pública cierre 2025</b> (CSV) y el <b>SICOP del corte actual 2026</b> (CSV) "
    "para generar el Formato de Austeridad "
    "</div>",
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)
with col1:
    archivo_cp = st.file_uploader(
        "Cuenta Pública cierre 2025",
        type=["csv"],
        key="uploader_cp",
        help="CSV del cierre de Cuenta Pública 2025 descargado del SICOP.",
    )
with col2:
    archivo_sicop = st.file_uploader(
        "SICOP corte actual 2026",
        type=["csv"],
        key="uploader_sicop",
        help="CSV del SICOP del corte del mes actual. Nombre sugerido: DD-MES-AAAA_SICOP.csv",
    )

ambos_listos = archivo_cp is not None and archivo_sicop is not None

if ambos_listos:
    st.caption(f"✔ {archivo_cp.name}   ✔ {archivo_sicop.name}")

if st.button(
    "Generar Formato de Austeridad",
    key="btn_aust",
    disabled=not ambos_listos,
):
    with st.spinner("Procesando…"):
        try:
            # Guardar temporales conservando el nombre original (para detectar corte)
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".csv", prefix="CP_"
            ) as t1:
                t1.write(archivo_cp.getvalue())
                ruta_cp = t1.name

            # Guardar SICOP con nombre original para detectar la fecha
            ruta_sicop_dir = tempfile.mkdtemp()
            ruta_sicop = os.path.join(ruta_sicop_dir, archivo_sicop.name)
            with open(ruta_sicop, "wb") as f:
                f.write(archivo_sicop.getvalue())

            # Generar el formato en memoria
            ruta_tmp_out = os.path.join(ruta_sicop_dir, "AUSTERIDAD_tmp.xlsx")
            ruta_out = generar_austeridad(ruta_cp, ruta_sicop, ruta_salida=ruta_tmp_out)

            with open(ruta_out, "rb") as f:
                buf_aust = io.BytesIO(f.read())
            buf_aust.seek(0)

            # Nombre de salida desde el corte detectado
            corte = detectar_corte(ruta_sicop)
            nombre_aust = (f"FORMATO_AUSTERIDAD_"
                           f"{corte['abr3']}-{str(corte['anio'])[2:]}.xlsx")

            # Limpiar temporales
            os.unlink(ruta_cp)
            os.unlink(ruta_sicop)
            os.unlink(ruta_out)
            os.rmdir(ruta_sicop_dir)

            st.success(
                f" Formato generado — corte al "
                f"{corte['dia']} de {corte['mes_es']} de {corte['anio']}"
            )
            st.download_button(
                label=f"⬇️  Descargar {nombre_aust}",
                data=buf_aust,
                file_name=nombre_aust,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_aust",
            )

        except Exception as e:
            st.error(f"Error al procesar: {e}")
            st.exception(e)
