import streamlit as st
import io, os, sys, tempfile

sys.path.insert(0, os.path.dirname(__file__))

from generar_cuadros_cocodi import (
    detectar_fecha, textos_fecha, leer_csv_map, leer_xlsx_map,
    hoja_pp, hoja_pp_cap, hoja_agricultura, hoja_comisario,
    MESES_ES, TPL_PP, TPL_PPCAP, TPL_AGR, TPL_COM, DIAS_FIN_MES,
)
from generar_austeridad import generar_austeridad, detectar_corte
import openpyxl

st.set_page_config(page_title="Reportes COCODI — SADER", page_icon="📊", layout="centered")

st.markdown("""
<style>
    .block-container { max-width: 740px; padding-top: 2rem; }
    .stDownloadButton > button {
        background-color: #9F2241; color: white; font-weight: 600;
        border: none; width: 100%; padding: 0.6rem 1.2rem;
        border-radius: 6px; font-size: 1rem;
    }
    .stDownloadButton > button:hover { background-color: #7a1a32; }
    .stButton > button {
        background-color: #BC955C; color: white; font-weight: 600;
        border: none; width: 100%; padding: 0.6rem 1.2rem;
        border-radius: 6px; font-size: 1rem;
    }
    .stButton > button:hover { background-color: #9a7840; }
    h1 { color: #9F2241; } h2 { color: #9F2241; margin-top: 0; }
    .info-box {
        background-color: #f9f3eb; border-left: 4px solid #BC955C;
        padding: 0.8rem 1rem; border-radius: 4px; margin-bottom: 1rem;
        font-size: 0.9rem; color: #444;
    }
    .section-divider { border: none; border-top: 2px solid #e8d5b0; margin: 2.5rem 0 2rem 0; }
</style>
""", unsafe_allow_html=True)

st.title("Reportes COCODI — SADER")

# ── SECCIÓN 1: Cuadros COCODI ─────────────────────────────────────────
st.markdown("##  Generador de COCODI")
st.markdown(
    "<div class='info-box'>Sube el archivo MAP del día (<b>.csv</b> o <b>.xlsx</b>) "
    "y descarga el Excel con las 4 hojas: <b>Pp</b>, <b>Pp y CAP</b>, "
    "<b>AGRICULTURA</b> y <b>Presupuesto Comisario</b>.</div>",
    unsafe_allow_html=True,
)

archivo_map = st.file_uploader("Archivo MAP del día", type=["csv","xlsx"], key="up_cocodi",
    help="Nombre sugerido: DD-MES-AAAA_MAP.csv")

if archivo_map:
    fecha_auto = detectar_fecha(archivo_map.name)
    st.markdown("**Verifica el periodo de corte** (detectado del nombre del archivo — ajústalo si no es correcto):")
    col1, col2, col3 = st.columns(3)
    with col1:
        anio_sel = st.number_input("Año", min_value=2020, max_value=2100,
                                    value=fecha_auto["anio"], key="anio_cocodi", step=1)
    with col2:
        mes_sel = st.selectbox("Mes de corte", options=list(range(1, 13)),
                                format_func=lambda m: MESES_ES[m].capitalize(),
                                index=fecha_auto["mes"] - 1, key="mes_cocodi")
    with col3:
        dia_sel = st.number_input("Día de corte", min_value=1, max_value=DIAS_FIN_MES[mes_sel],
                                   value=min(fecha_auto["dia"], DIAS_FIN_MES[mes_sel]),
                                   key="dia_cocodi", step=1)
    st.caption(f"Se generará el cuadro con corte al {dia_sel} de {MESES_ES[mes_sel]} de {anio_sel}")

    if st.button("Generar cuadros COCODI", key="btn_cocodi"):
        with st.spinner("Procesando…"):
            try:
                ext = os.path.splitext(archivo_map.name)[1].lower()
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="MAP_") as tmp:
                    tmp.write(archivo_map.getvalue()); ruta_tmp = tmp.name
                fecha = {"dia": dia_sel, "mes": mes_sel, "anio": anio_sel}
                tf = textos_fecha(fecha)
                datos = leer_csv_map(ruta_tmp, fecha["mes"]) if ext==".csv" else leer_xlsx_map(ruta_tmp)
                datos["_mes"] = fecha["mes"]
                tpls = {"Pp":TPL_PP,"PpCAP":TPL_PPCAP,"AGRICULTURA":TPL_AGR,"Comisario":TPL_COM}
                wb_out = openpyxl.Workbook()
                if "Sheet" in wb_out.sheetnames: del wb_out["Sheet"]
                hoja_pp(wb_out,datos,tf,tpls["Pp"])
                hoja_pp_cap(wb_out,datos,tf,tpls["PpCAP"])
                hoja_agricultura(wb_out,datos,tf,tpls["AGRICULTURA"])
                hoja_comisario(wb_out,datos,tf,tpls["Comisario"])
                buf = io.BytesIO(); wb_out.save(buf); buf.seek(0)
                mes_str = MESES_ES[fecha["mes"]].upper()
                nombre_out = f"CUADROS_COCODI_{fecha['dia']:02d}-{mes_str}-{fecha['anio']}.xlsx"
                os.unlink(ruta_tmp)
                st.success(f" Cuadros generados — corte al {tf['dia_mes_anio']}")
                st.download_button(label=f"⬇️  Descargar {nombre_out}", data=buf,
                    file_name=nombre_out, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_cocodi")
            except Exception as e:
                st.error(f"Error: {e}"); st.exception(e)

# ── SECCIÓN 2: Formato de Austeridad ─────────────────────────────────
st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
st.markdown("##  Formato de Austeridad")
st.markdown(
    "<div class='info-box'>"
    "Sube el <b>SICOP del corte actual 2026</b> (CSV) para generar el Formato de Austeridad. "
    "Los valores de la Cuenta Pública 2025 ya están incorporados — "
    "solo necesitas el SICOP del mes actual."
    "</div>",
    unsafe_allow_html=True,
)

archivo_sicop = st.file_uploader("SICOP corte actual 2026", type=["csv"], key="up_sicop",
    help="Nombre sugerido: DD-MES-AAAA_SICOP.csv")

if archivo_sicop:
    st.caption(f"✔ {archivo_sicop.name}")
    if st.button("Generar Formato de Austeridad", key="btn_aust"):
        with st.spinner("Procesando…"):
            try:
                tmpdir = tempfile.mkdtemp()
                ruta_sicop = os.path.join(tmpdir, archivo_sicop.name)
                with open(ruta_sicop,"wb") as f: f.write(archivo_sicop.getvalue())
                ruta_out = os.path.join(tmpdir, "AUSTERIDAD_tmp.xlsx")
                generar_austeridad(ruta_sicop, ruta_salida=ruta_out)
                with open(ruta_out,"rb") as f: buf = io.BytesIO(f.read()); buf.seek(0)
                corte = detectar_corte(ruta_sicop)
                nombre_aust = f"FORMATO_AUSTERIDAD_{corte['abr3']}-{str(corte['anio'])[2:]}.xlsx"
                os.unlink(ruta_sicop); os.unlink(ruta_out); os.rmdir(tmpdir)
                st.success(f" Formato generado — corte al {corte['dia']} de {corte['mes_es']} de {corte['anio']}")
                st.download_button(label=f"⬇️  Descargar {nombre_aust}", data=buf,
                    file_name=nombre_aust, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_aust")
            except Exception as e:
                st.error(f"Error: {e}"); st.exception(e)
