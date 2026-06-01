import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")


# =========================
# FUNCIONES
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
# DATOS GLOBALES (1 SOLO VEZ)
# =========================

st.header("📂 Carga de datos base")

datos_file = st.file_uploader("Sube datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("Sube Lista_Espera.xlsx", type=["xlsx"])

# =========================
# VALIDACIÓN GLOBAL
# =========================

if not (datos_file and lp_file):
    st.info("📌 Sube datos y lista de espera para habilitar los informes")

else:

    st.header("📑 Informes Preliminares")

    col1, col2 = st.columns(2)

    # -------------------------
    # PRELIMINAR 1
    # -------------------------
    with col1:

        st.subheader("Preliminar 1")

        word1 = st.file_uploader(
            "Plantilla Preliminar 1",
            type=["docx"],
            key="w1"
        )

        if st.button("Generar Preliminar 1"):

            doc = DocxTemplate(word1)

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

    # -------------------------
    # PRELIMINAR 2
    # -------------------------
    with col2:

        st.subheader("Preliminar 2")

        word2 = st.file_uploader(
            "Plantilla Preliminar 2",
            type=["docx"],
            key="w2"
        )

        if st.button("Generar Preliminar 2"):

            doc = DocxTemplate(word2)

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
