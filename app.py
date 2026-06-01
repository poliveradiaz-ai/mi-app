import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")

st.write("Sube los archivos base y opcionalmente las plantillas de informes")

# =========================
# ARCHIVOS BASE (OBLIGATORIOS)
# =========================
st.markdown("## 📂 Archivos base (obligatorios)")

datos_file = st.file_uploader("📄 Subir datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("📄 Subir Lista_Espera.xlsx", type=["xlsx"])

st.divider()

# =========================
# PLANTILLAS (OPCIONALES)
# =========================
st.markdown("## 🧾 Plantillas de informes (opcionales)")

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
    word_file_1 = st.file_uploader(
        "Subir plantilla 1 (.docx)",
        type=["docx"],
        key="w1"
    )

with col2:
    st.markdown(
        """
        <div style='padding:15px; border-radius:12px; background:#f1fff3; border-left:6px solid #2ecc71;'>
        <h4>📄 Informe Preliminar 2</h4>
        </div>
        """,
        unsafe_allow_html=True
    )
    word_file_2 = st.file_uploader(
        "Subir plantilla 2 (.docx)",
        type=["docx"],
        key="w2"
    )

# =========================
# BOTÓN
# =========================
if st.button("🚀 Generar Reportes"):

    if datos_file and lp_file:

        try:
            # =========================
            # LECTURA DE ARCHIVOS
            # =========================
            df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")

            # =========================
            # PLANTILLAS (opcionales)
            # =========================
            doc1 = DocxTemplate(word_file_1) if word_file_1 else None
            doc2 = DocxTemplate(word_file_2) if word_file_2 else None

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
                lp[['rut_puente', 'esp_puente', 'Num Interconsulta']],
                on=['rut_puente', 'esp_puente'],
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

            # =========================
            # RENDER (solo si existen plantillas)
            # =========================
            if doc1:
                doc1.render(contexto)

            if doc2:
                doc2.render(contexto)

            # =========================
            # GUARDADO
            # =========================
            st.success("✅ Proceso completado correctamente")

            colA, colB = st.columns(2)

            # -------------------------
            # INFORME 1
            # -------------------------
            if doc1:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                    doc1.save(tmp1.name)

                with open(tmp1.name, "rb") as f1:
                    colA.download_button(
                        "📥 Descargar Informe 1",
                        f1,
                        file_name="Informe_Preliminar_1.docx"
                    )
            else:
                colA.warning("⚠️ No subiste plantilla 1")

            # -------------------------
            # INFORME 2
            # -------------------------
            if doc2:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                    doc2.save(tmp2.name)

                with open(tmp2.name, "rb") as f2:
                    colB.download_button(
                        "📥 Descargar Informe 2",
                        f2,
                        file_name="Informe_Preliminar_2.docx"
                    )
            else:
                colB.warning("⚠️ No subiste plantilla 2")

            # =========================
            # KPIs
            # =========================
            st.markdown("## 📊 Resumen")

            c1, c2, c3, c4 = st.columns(4)

            c1.metric("Control", total_escontrol)
            c2.metric("CN", total_escn)
            c3.metric("Interconsultas", total_inter)
            c4.metric("Resultado", resultado)

        except Exception as e:
            st.error(f"Error: {e}")

    else:
        st.warning("⚠️ Debes subir al menos los archivos base (datos y lista de espera)")
