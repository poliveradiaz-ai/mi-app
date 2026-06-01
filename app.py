import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")

st.write("Sube los archivos base y las plantillas para generar los informes")

# =========================
# ARCHIVOS BASE (COMUNES)
# =========================
st.markdown("## 📂 Archivos base (comunes para ambos informes)")

datos_file = st.file_uploader("📄 Subir datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("📄 Subir Lista_Espera.xlsx", type=["xlsx"])

st.divider()

# =========================
# TARJETAS DE PLANTILLAS
# =========================
st.markdown("## 🧾 Plantillas de informes")

col1, col2 = st.columns(2)

with col1:
    st.markdown(
        """
        <div style='padding:15px; border-radius:12px; background:#f1f8ff; border-left:6px solid #3498db;'>
        <h4>📄 Informe Preliminar 1</h4>
        </div>
        """,
        unsafe_allow_html=True
    )
    word_file_1 = st.file_uploader("Subir plantilla 1 (.docx)", type=["docx"], key="w1")

with col2:
    st.markdown(
        """
        <div style='padding:15px; border-radius:12px; background:#f1fff3; border-left:6px solid #2ecc71;'>
        <h4>📄 Informe Preliminar 2</h4>
        </div>
        """,
        unsafe_allow_html=True
    )
    word_file_2 = st.file_uploader("Subir plantilla 2 (.docx)", type=["docx"], key="w2")


# =========================
# BOTÓN
# =========================
if st.button("🚀 Generar Reportes"):

    if datos_file and lp_file and word_file_1 and word_file_2:

        try:
            # =========================
            # LEER DATOS
            # =========================
            df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")

            doc1 = DocxTemplate(word_file_1)
            doc2 = DocxTemplate(word_file_2)

            # =========================
            # PROCESAMIENTO BASE
            # =========================
            actividad = df['Actividad'].astype(str).str.strip().str.upper()

            total_controles = (actividad == 'CONTROL').sum()
            total_consultas_nuevas = (actividad == 'CONSULTA NUEVA').sum()

            df['Funcionario'] = (
                df['Nombres'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Pat'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Mat'].fillna('').astype(str).str.strip()
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

            df['rut_puente'] = df['Rut'].astype(str).str.replace(".", "").str.split("-").str[0]
            lp['rut_puente'] = lp['Rut'].astype(str).str.replace(".", "").str.split("-").str[0]

            df['esp_puente'] = df['Especialidad'].astype(str).str.strip().str.upper()
            lp['esp_puente'] = lp['Especialidad Destino'].astype(str).str.strip().str.upper()

            df = pd.merge(
                df,
                lp[['rut_puente','esp_puente','Num Interconsulta']],
                on=['rut_puente','esp_puente'],
                how='left'
            )

            df['Num Interconsulta'] = df['Num Interconsulta'].fillna(0)

            # =========================
            # CÁLCULOS
            # =========================
            total_escontrol = (actividad == 'CONTROL').sum()
            total_escn = (actividad == 'CONSULTA NUEVA').sum()
            total_inter = (df['Num Interconsulta'] > 0).sum()

            resultado = total_escontrol - total_inter

            # =========================
            # CONTEXTO WORD
            # =========================
            contexto = {
                "total_es_control": int(total_escontrol),
                "total_es_cn": int(total_escn),
                "total_controles": int(total_controles),
                "total_consultas_nuevas": int(total_consultas_nuevas),
                "total_inter": int(total_inter),
                "resultado": int(resultado),
                "filas": df.to_dict("records")
            }

            doc1.render(contexto)
            doc2.render(contexto)

            # =========================
            # GUARDAR
            # =========================
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                doc1.save(tmp1.name)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                doc2.save(tmp2.name)

            st.success("✅ Reportes generados correctamente")

            colA, colB = st.columns(2)

            with open(tmp1.name, "rb") as f1:
                colA.download_button(
                    "📥 Descargar Informe 1",
                    f1,
                    file_name="Informe_Preliminar_1.docx"
                )

            with open(tmp2.name, "rb") as f2:
                colB.download_button(
                    "📥 Descargar Informe 2",
                    f2,
                    file_name="Informe_Preliminar_2.docx"
                )

            # =========================
            # DASHBOARD FINAL
            # =========================
            st.markdown("## 📊 Resumen General")

            c1, c2, c3, c4 = st.columns(4)

            c1.metric("Control", total_escontrol)
            c2.metric("CN", total_escn)
            c3.metric("Interconsultas", total_inter)
            c4.metric("Resultado", resultado)

        except Exception as e:
            st.error(f"Error: {e}")

    else:
        st.warning("⚠️ Debes subir todos los archivos antes de continuar")
