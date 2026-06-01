import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")


# =========================
# FUNCIONES BASE
# =========================

def calcular_controlycn(fila):
    actividad = str(fila.get('Actividad','')).strip().upper()
    Ic = str(fila.get('Ic Asoc Hora','')).strip()

    rEs_control = 1 if (actividad == 'CONSULTA NUEVA' and Ic == '-') else 0
    rEs_CN = 1 if (actividad == 'CONTROL' and Ic != '-') else 0

    return pd.Series([rEs_control, rEs_CN])


def limpiar_rut_definitivo(rut):
    rut = str(rut).strip()
    if "-" in rut:
        rut = rut.split("-")[0]
    return rut.replace(".", "")


def marcar_interconsulta_valida(fila):
    actividad = str(fila.get('Actividad', '')).strip().upper()
    ic_asoc = str(fila.get('Ic Asoc Hora', '')).strip()
    num_ic = fila.get('Num Interconsulta', 0)

    try:
        num_ic = float(num_ic)
    except:
        num_ic = 0

    return 1 if (
        actividad == 'CONSULTA NUEVA'
        and ic_asoc == '-'
        and num_ic != 0
    ) else 0


# =========================
# PRELIMINAR 1 (WORD + LISTA ESPERA)
# =========================

def generar_preliminar_1(datos_file, lp_file, word_file):

    df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
    lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")
    doc = DocxTemplate(word_file)

    actividad = df['Actividad'].astype(str).str.strip().str.upper()

    total_controles = (actividad == 'CONTROL').sum()
    total_consultas_nuevas = (actividad == 'CONSULTA NUEVA').sum()

    df[['Es_control', 'Es_CN']] = df.apply(calcular_controlycn, axis=1)

    total_escontrol = int(df['Es_control'].sum())

    df['rut_puente'] = df['Rut'].apply(limpiar_rut_definitivo)
    lp['rut_puente'] = lp['Rut'].apply(limpiar_rut_definitivo)

    df['esp_puente'] = df['Especialidad'].str.strip().str.upper()
    lp['esp_puente'] = lp['Especialidad Destino'].str.strip().str.upper()

    df = pd.merge(
        df,
        lp[['rut_puente','esp_puente','Num Interconsulta']],
        on=['rut_puente','esp_puente'],
        how='left'
    )

    df['Num Interconsulta'] = df['Num Interconsulta'].fillna(0)

    df['Interconsulta_Valida'] = df.apply(marcar_interconsulta_valida, axis=1)

    total_inter = int(df['Interconsulta_Valida'].sum())

    resultado = total_escontrol - total_inter

    porc_1 = (resultado / total_controles * 100) if total_controles > 0 else 0
    porc_2 = (resultado / total_consultas_nuevas * 100) if total_consultas_nuevas > 0 else 0

    contexto = {
        "total_escontrol": total_escontrol,
        "total_inter": total_inter,
        "resultado": resultado,
        "porc_controles": round(porc_1, 2),
        "porc_cn": round(porc_2, 2),
    }

    doc.render(contexto)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button(
                "📥 Descargar Preliminar 1",
                f,
                file_name="Preliminar_1.docx"
            )

    st.success("Preliminar 1 generado")


# =========================
# PRELIMINAR 2 (WORD + SIN LISTA ESPERA)
# =========================

def generar_preliminar_2(datos_file, word_file):

    df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
    doc = DocxTemplate(word_file)

    actividad = df['Actividad'].astype(str).str.strip().str.upper()

    total_controles = (actividad == 'CONTROL').sum()
    total_consultas_nuevas = (actividad == 'CONSULTA NUEVA').sum()

    df[['Es_control', 'Es_CN']] = df.apply(calcular_controlycn, axis=1)

    total_cn_error = int(df['Es_CN'].sum())

    porc_vs_controles = (
        (total_cn_error / total_controles) * 100
        if total_controles > 0 else 0
    )

    porc_vs_cn = (
        (total_cn_error / total_consultas_nuevas) * 100
        if total_consultas_nuevas > 0 else 0
    )

    contexto = {
        "total_cn_error": total_cn_error,
        "porc_controles": round(porc_vs_controles, 2),
        "porc_cn": round(porc_vs_cn, 2),
    }

    doc.render(contexto)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button(
                "📥 Descargar Preliminar 2",
                f,
                file_name="Preliminar_2.docx"
            )

    st.success("Preliminar 2 generado")


# =========================
# INTERFAZ
# =========================

col1, col2 = st.columns(2)

# -------------------------
# PRELIMINAR 1
# -------------------------
with col1:

    st.subheader("📄 Preliminar 1")

    datos1 = st.file_uploader("Datos", type=["xlsx"], key="d1")
    lp1 = st.file_uploader("Lista Espera", type=["xlsx"], key="lp1")
    word1 = st.file_uploader("Plantilla Word 1", type=["docx"], key="w1")

    if st.button("Generar Preliminar 1", key="b1"):
        generar_preliminar_1(datos1, lp1, word1)


# -------------------------
# PRELIMINAR 2
# -------------------------
with col2:

    st.subheader("📄 Preliminar 2")

    datos2 = st.file_uploader("Datos", type=["xlsx"], key="d2")
    word2 = st.file_uploader("Plantilla Word 2", type=["docx"], key="w2")

    if st.button("Generar Preliminar 2", key="b2"):
        generar_preliminar_2(datos2, word2)
