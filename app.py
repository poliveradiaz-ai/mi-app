import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile
from io import BytesIO

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")

st.write("Sube los archivos base y opcionalmente las plantillas")

# =========================
# ARCHIVOS BASE
# =========================
st.markdown("## 📂 Archivos base")

datos_file = st.file_uploader("📄 Subir datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("📄 Subir Lista_Espera.xlsx", type=["xlsx"])

st.divider()

# =========================
# FUNCIONES
# =========================
def calcular_controlycn(fila):
    actividad = str(fila.get('Actividad','')).strip()
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

    return 1 if (actividad == 'CONSULTA NUEVA' and ic_asoc == '-' and num_ic != 0) else 0


# =========================
# PLANTILLAS
# =========================
col1, col2 = st.columns(2)

with col1:
    st.markdown("📄 Informe Preliminar 1")
    word_file = st.file_uploader("Plantilla 1 (.docx)", type=["docx"], key="w1")

with col2:
    st.markdown("📄 Informe Preliminar 2")
    preliminar2_word_file = st.file_uploader("Plantilla 2 (.docx)", type=["docx"], key="w2")


# =========================
# INIT SESSION STATE
# =========================
if "informe1" not in st.session_state:
    st.session_state["informe1"] = None

if "informe2" not in st.session_state:
    st.session_state["informe2"] = None


# =========================
# BOTÓN GENERAR
# =========================
if st.button("🚀 Generar Reporte"):

    if datos_file and lp_file:

        try:
            df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")

            doc = DocxTemplate(word_file) if word_file else None
            doc2 = DocxTemplate(preliminar2_word_file) if preliminar2_word_file else None

            actividad = (
                df['Actividad']
                .astype(str)
                .str.strip()
                .str.upper()
            )

            total_controles = (actividad == 'CONTROL').sum()
            total_consultas_nuevas = (actividad == 'CONSULTA NUEVA').sum()

            df[['Es_control', 'Es_CN']] = df.apply(calcular_controlycn, axis=1)

            total_escontrol = int(df['Es_control'].sum())
            total_escn = int(df['Es_CN'].sum())

            df['Funcionario'] = (
                df['Nombres'].fillna('').astype(str) + ' ' +
                df['Apellido Pat'].fillna('').astype(str) + ' ' +
                df['Apellido Mat'].fillna('').astype(str)
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

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

            df['Error_escontrol_menos_Inter'] = (
                (df['Es_control'] - df['Interconsulta_Valida']) > 0
            ).astype(int)

            df['Especialidad'] = (
                df['Especialidad']
                .astype(str)
                .str.strip()
                .str.upper()
                .replace({
                    # TRAUMATOLOGIA
                    'TRAUMATOLOGIA Y ORTOPEDIA ADULTO': 'TRAUMATOLOGIA Y ORTOPEDIA',
                    'TRAUMATOLOGÍA Y ORTOPEDIA ADULTO': 'TRAUMATOLOGIA Y ORTOPEDIA',
                    'TRAUMATOLOGÍA Y ORTOPEDIA': 'TRAUMATOLOGIA Y ORTOPEDIA',
           
                    # GINECOLOGIA
                    'GINECOLOGIA GENERAL ADULTO': 'GINECOLOGIA',
                    'GINECOLOGÍA GENERAL ADULTO': 'GINECOLOGIA',
                    'GINECOLOGÍA': 'GINECOLOGIA'
                })
            )

            df_controles = df[df['Error_escontrol_menos_Inter'] == 1]
            df_escn = df[df['Es_CN'] == 1]

            df_controles = df[df['Error_escontrol_menos_Inter'] == 1].copy()
            df_escn = df[df['Es_CN'] == 1].copy()
           
            tabla = (
                df_controles
                .groupby("Especialidad")
                .size()
                .reset_index(name="cantidad")
                .sort_values(by="cantidad", ascending=False)
            )

            orden_especialidades = tabla['Especialidad'].tolist()

            tabla_funcionarios = (
                df_controles
                .groupby(['Especialidad', 'Funcionario'])
                .size()
                .reset_index(name='total')
            )
           
            tabla_funcionarios['orden_esp'] = tabla_funcionarios['Especialidad'].map(
                {esp: i for i, esp in enumerate(orden_especialidades)}
            )
           
            tabla_funcionarios = tabla_funcionarios.sort_values(
                by=['orden_esp', 'total'],
                ascending=[True, False]
            ).drop(columns=['orden_esp'])
           
            tabla_escn = (
                df_escn
                .groupby('Especialidad')
                .size()
                .reset_index(name='cantidad')
                .sort_values(by='cantidad', ascending=False)
            )

            orden_especialidades_cn = (
                tabla_escn
                .sort_values(by='cantidad', ascending=False)
                ['Especialidad']
                .tolist()
            )
           
            tabla_funcionarios_cn = (
                df_escn
                .groupby(['Especialidad', 'Funcionario'])
                .size()
                .reset_index(name='total')
            )
           
            tabla_funcionarios_cn['orden_esp'] = tabla_funcionarios_cn['Especialidad'].map(
                {esp: i for i, esp in enumerate(orden_especialidades_cn)}
            )
           
            tabla_funcionarios_cn = tabla_funcionarios_cn.sort_values(
                by=['orden_esp', 'total'],
                ascending=[True, False]
            ).drop(columns=['orden_esp'])

           
            contexto = {
                'filas': df.to_dict('records'),
                'total_es_control': total_escontrol,
                'total_es_cn': total_escn,
                'total_controles': total_controles,
                'total_consultas_nuevas': total_consultas_nuevas,
                'total_inter': total_inter,
                'resultado_es_control_menos_interconsulta': resultado,
                'especialidades': tabla.to_dict('records'),
                'tabla_funcionarios': tabla_funcionarios.to_dict('records'),
                'tabla_escn': tabla_escn.to_dict('records'),
                'tabla_funcionarios_cn': tabla_funcionarios_cn.to_dict('records'),
            }

            if doc:
                doc.render(contexto)
                buffer1 = BytesIO()
                doc.save(buffer1)
                st.session_state["informe1"] = buffer1.getvalue()

            if doc2:
                doc2.render(contexto)
                buffer2 = BytesIO()
                doc2.save(buffer2)
                st.session_state["informe2"] = buffer2.getvalue()

            st.success("✅ Reporte generado correctamente")

        import traceback

    except Exception as e:
        st.error(f"Error: {e}")
        st.code(traceback.format_exc())

    else:
        st.warning("Debes subir los archivos base")


# =========================
# DESCARGAS (SIEMPRE ESTABLES)
# =========================
st.divider()
st.subheader("📥 Descargar Informes")

colA, colB = st.columns(2)

with colA:
    if st.session_state["informe1"]:
        st.download_button(
            "📥 Informe 1",
            data=st.session_state["informe1"],
            file_name="Informe_1.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_1"
        )

with colB:
    if st.session_state["informe2"]:
        st.download_button(
            "📥 Informe 2",
            data=st.session_state["informe2"],
            file_name="Informe_2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_2"
        )
