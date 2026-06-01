import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")

st.write("Sube los archivos base y opcionalmente las plantillas de informes")

# =========================
# ARCHIVOS BASE
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
    st.markdown("### 📄 Informe Preliminar 1")
    word_file_1 = st.file_uploader("Subir plantilla 1", type=["docx"], key="w1")

with col2:
    st.markdown("### 📄 Informe Preliminar 2")
    word_file_2 = st.file_uploader("Subir plantilla 2", type=["docx"], key="w2")


# =========================
# BOTÓN
# =========================
if st.button("🚀 Generar Reportes"):

    if datos_file and lp_file:

        try:
            # =========================
            # LECTURA
            # =========================
            df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")

            doc1 = DocxTemplate(word_file_1) if word_file_1 else None
            doc2 = DocxTemplate(word_file_2) if word_file_2 else None

            # =========================
            # ACTIVIDAD
            # =========================
            actividad = df['Actividad'].astype(str).str.strip().str.upper()

            total_controles = (actividad == 'CONTROL').sum()
            total_consultas_nuevas = (actividad == 'CONSULTA NUEVA').sum()

            # =========================
            # FUNCIONARIO
            # =========================
            df['Funcionario'] = (
                df['Nombres'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Pat'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Mat'].fillna('').astype(str).str.strip()
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

            # =========================
            # RUT MERGE
            # =========================
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
            # ES_CONTROL / ES_CN
            # =========================
            df['Es_control'] = ((actividad == 'CONSULTA NUEVA') & (df['Ic Asoc Hora'].astype(str).str.strip() == '-')).astype(int)
            df['Es_CN'] = ((actividad == 'CONTROL') & (df['Ic Asoc Hora'].astype(str).str.strip() != '-')).astype(int)

            total_escontrol = df['Es_control'].sum()
            total_escn = df['Es_CN'].sum()

            # =========================
            # INTERCONSULTA VÁLIDA
            # =========================
            df['Interconsulta_Valida'] = (
                (actividad == 'CONSULTA NUEVA') &
                (df['Ic Asoc Hora'].astype(str).str.strip() == '-') &
                (df['Num Interconsulta'] != 0)
            ).astype(int)

            total_inter = df['Interconsulta_Valida'].sum()

            resultado = total_escontrol - total_inter

            # =========================
            # ERROR (CLAVE CORRECTA)
            # =========================
            df['Error_escontrol_menos_Inter'] = (
                (df['Es_control'] - df['Interconsulta_Valida']) > 0
            ).astype(int)

            # 🔥 FILTRO CORRECTO (ESTO TE ESTABA FALLANDO)
            df_controles = df[df['Error_escontrol_menos_Inter'] == 1]

            # =========================
            # TABLAS CORRECTAS
            # =========================
            tabla = (
                df_controles
                .groupby("Especialidad")
                .size()
                .reset_index(name="cantidad")
                .sort_values(by="cantidad", ascending=False)
            )

            tabla_funcionarios = (
                df_controles
                .groupby(['Especialidad', 'Funcionario'])
                .size()
                .reset_index(name='total')
                .sort_values(by='total', ascending=False)
            )

            # =========================
            # CONTEXTO WORD (CORRECTO)
            # =========================
            contexto = {
                'filas': df.to_dict('records'),

                'total_es_control': int(total_escontrol),
                'total_es_cn': int(total_escn),

                'total_controles': int(total_controles),
                'total_consultas_nuevas': int(total_consultas_nuevas),

                'total_inter': int(total_inter),
                'resultado_es_control_menos_interconsulta': int(resultado),

                'especialidades': tabla.to_dict('records'),
                'tabla_funcionarios': tabla_funcionarios.to_dict('records'),
                'total_general_control': int(tabla['cantidad'].sum()),
            }

            # =========================
            # RENDER WORD
            # =========================
            if doc1:
                doc1.render(contexto)

            if doc2:
                doc2.render(contexto)

            # =========================
            # GUARDADO
            # =========================
            st.success("✅ Reportes generados correctamente")

            colA, colB = st.columns(2)

            if doc1:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                    doc1.save(tmp1.name)

                with open(tmp1.name, "rb") as f1:
                    colA.download_button(
                        "📥 Informe 1",
                        f1,
                        file_name="Informe_1.docx"
                    )
            else:
                colA.warning("No subiste plantilla 1")

            if doc2:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                    doc2.save(tmp2.name)

                with open(tmp2.name, "rb") as f2:
                    colB.download_button(
                        "📥 Informe 2",
                        f2,
                        file_name="Informe_2.docx"
                    )
            else:
                colB.warning("No subiste plantilla 2")

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
        st.warning("⚠️ Debes subir datos y lista de espera")
