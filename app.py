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
# CARGA DE ARCHIVOS
# =========================

st.header("📂 Carga de datos base")

datos_file = st.file_uploader("Sube datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("Sube Lista_Espera.xlsx", type=["xlsx"])


df_base = None
lp_base = None


# =========================
# PREPROCESAMIENTO GLOBAL
# =========================

if datos_file and lp_file:

    df_base = pd.read_excel(
        datos_file,
        sheet_name="NOMINA CUADRATURA (REM7) SIN CO"
    )

    lp_base = pd.read_excel(
        lp_file,
        sheet_name="Nomina Médico"
    )

    # =========================
    # FLAGS
    # =========================
    df_base[['Es_control', 'Es_CN']] = df_base.apply(
        calcular_controlycn,
        axis=1
    )

    actividad = df_base['Actividad'].astype(str).str.strip().str.upper()

    total_controles = (actividad == 'CONTROL').sum()
    total_consultas_nuevas = (actividad == 'CONSULTA NUEVA').sum()

    total_escontrol = int(df_base['Es_control'].sum())
    total_cn_error = int(df_base['Es_CN'].sum())

    # =========================
    # RUT + MERGE
    # =========================
    df_base['rut_puente'] = df_base['Rut'].apply(limpiar_rut_definitivo)
    lp_base['rut_puente'] = lp_base['Rut'].apply(limpiar_rut_definitivo)

    df_base['esp_puente'] = df_base['Especialidad'].str.strip().str.upper()
    lp_base['esp_puente'] = lp_base['Especialidad Destino'].str.strip().str.upper()

    df_base = pd.merge(
        df_base,
        lp_base[['rut_puente','esp_puente','Num Interconsulta']],
        on=['rut_puente','esp_puente'],
        how='left'
    )

    df_base['Num Interconsulta'] = df_base['Num Interconsulta'].fillna(0)

    # =========================
    # INTERCONSULTA VALIDA
    # =========================
    df_base['Interconsulta_Valida'] = df_base.apply(
        marcar_interconsulta_valida,
        axis=1
    )

    total_inter = int(df_base['Interconsulta_Valida'].sum())

    # =========================
    # 🔴 FIX CRÍTICO: ERROR_FINAL (ESTO TE FALTABA)
    # =========================
    df_base['Error_Final'] = (
        df_base['Es_control'] - df_base['Interconsulta_Valida']
    )

    df_base['Error_Final'] = (df_base['Error_Final'] > 0).astype(int)

    resultado = total_escontrol - total_inter


# =========================
# INTERFAZ
# =========================

st.header("📑 Informes Preliminares")

col1, col2 = st.columns(2)


# =========================
# PRELIMINAR 1
# =========================

with col1:

    st.subheader("Preliminar 1")

    word1 = st.file_uploader(
        "Plantilla Preliminar 1",
        type=["docx"],
        key="w1"
    )

    if st.button("Generar Preliminar 1"):

        if df_base is not None:

            doc = DocxTemplate(word1)

            contexto = {
                'filas': df_base.to_dict(orient='records'),

                'total_es_control': total_escontrol,
                'total_es_cn': total_cn_error,

                'total_controles': int(total_controles),
                'total_consultas_nuevas': int(total_consultas_nuevas),

                'total_inter': total_inter,

                'resultado_es_control_menos_interconsulta': resultado,

                'porc_escontrol_vs_controles': round(
                    (resultado / total_controles * 100) if total_controles > 0 else 0, 2
                ),

                'porc_escontrol_vs_cn': round(
                    (resultado / total_consultas_nuevas * 100) if total_consultas_nuevas > 0 else 0, 2
                ),

                'especialidades': df_base[df_base['Error_Final'] == 1]
                    .groupby("Especialidad")
                    .size()
                    .reset_index(name="cantidad")
                    .to_dict(orient='records'),

                'total_general_control': int(
                    df_base[df_base['Error_Final'] == 1]
                    .groupby("Especialidad")
                    .size()
                    .sum()
                ),

                'tabla_funcionarios': df_base[df_base['Error_Final'] == 1]
                    .groupby(['Especialidad', 'Funcionario'])
                    .size()
                    .reset_index(name='total')
                    .to_dict(orient='records'),
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

        else:
            st.warning("Sube los archivos primero")


# =========================
# PRELIMINAR 2
# =========================

with col2:

    st.subheader("Preliminar 2")

    word2 = st.file_uploader(
        "Plantilla Preliminar 2",
        type=["docx"],
        key="w2"
    )

    if st.button("Generar Preliminar 2"):

        if df_base is not None:

            doc = DocxTemplate(word2)

            contexto = {
                'total_es_control': total_escontrol,
                'total_es_cn': total_cn_error,

                'total_controles': int(total_controles),
                'total_consultas_nuevas': int(total_consultas_nuevas),

                'total_inter': total_inter,

                'resultado_es_control_menos_interconsulta': total_cn_error,

                'porc_escontrol_vs_controles': round(
                    (total_cn_error / total_controles * 100) if total_controles > 0 else 0, 2
                ),

                'porc_escontrol_vs_cn': round(
                    (total_cn_error / total_consultas_nuevas * 100) if total_consultas_nuevas > 0 else 0, 2
                ),
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

        else:
            st.warning("Sube los archivos primero")


# =========================
# MENSAJE FINAL
# =========================

if df_base is None:
    st.info("📌 Sube datos y lista de espera para habilitar los informes")
