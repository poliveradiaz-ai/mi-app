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
# FUNCIONES (ORIGINALES - OBLIGATORIAS)
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

    if actividad == 'CONSULTA NUEVA' and ic_asoc == '-' and num_ic != 0:
        return 1
    return 0


# =========================
# TARJETAS PLANTILLAS
# =========================
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div style="padding:15px;border-radius:12px;background:#e8f4ff;border-left:6px solid #1f77b4;">
        <h4>📄 Informe Preliminar 1</h4>
    </div>
    """, unsafe_allow_html=True)

    word_file = st.file_uploader("Plantilla 1 (.docx)", type=["docx"], key="w1")

with col2:
    st.markdown("""
    <div style="padding:15px;border-radius:12px;background:#eaffea;border-left:6px solid #2ca02c;">
        <h4>📄 Informe Preliminar 2</h4>
    </div>
    """, unsafe_allow_html=True)

    preliminar2_word_file = st.file_uploader("Plantilla 2 (.docx)", type=["docx"], key="w2")


# =========================
# BOTÓN
# =========================
if st.button("🚀 Generar Reporte"):

    if datos_file and lp_file:

        try:
            # =========================
            # LECTURA
            # =========================
            df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")

            doc = DocxTemplate(word_file) if word_file else None
            doc2 = DocxTemplate(preliminar2_word_file) if preliminar2_word_file else None

            # =========================
            # TOTALES (ORIGINAL)
            # =========================
            actividad = (
                df['Actividad']
                .astype(str)
                .str.strip()
                .str.upper()
            )

            total_controles = (actividad == 'CONTROL').sum()
            total_consultas_nuevas = (actividad == 'CONSULTA NUEVA').sum()

            # =========================
            # ES_CONTROL / ES_CN (ORIGINAL)
            # =========================
            df[['Es_control', 'Es_CN']] = df.apply(calcular_controlycn, axis=1)

            total_escontrol = int(df['Es_control'].sum())
            total_escn = int(df['Es_CN'].sum())

            # =========================
            # FUNCIONARIO (ORIGINAL)
            # =========================
            df['Funcionario'] = (
                df['Nombres'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Pat'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Mat'].fillna('').astype(str).str.strip()
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

            # =========================
            # RUT + MERGE (ORIGINAL)
            # =========================
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

            # =========================
            # INTERCONSULTA (ORIGINAL)
            # =========================
            df['Interconsulta_Valida'] = df.apply(marcar_interconsulta_valida, axis=1)

            total_inter = int(df['Interconsulta_Valida'].sum())

            resultado = total_escontrol - total_inter

            # =========================
            # ERROR FINAL (ORIGINAL)
            # =========================
            df['Error_escontrol_menos_Inter'] = (df['Es_control'] - df['Interconsulta_Valida'])
            df['Error_escontrol_menos_Inter'] = (df['Error_escontrol_menos_Inter'] > 0).astype(int)

            # =========================
            # NORMALIZAR ESPECIALIDADES (CORRECCIÓN PEDIDA)
            # =========================
            df['Especialidad'] = (
                df['Especialidad']
                .astype(str)
                .str.strip()
                .replace({
                    'TRAUMATOLOGIA Y ORTOPEDIA ADULTO': 'TRAUMATOLOGIA Y ORTOPEDIA'
                })
            )
            df_controles = df[df['Error_escontrol_menos_Inter'] == 1]
            df_escn = df[df['Es_CN'] == 1]
            # =========================
            # TABLAS (ORIGINAL)
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
  
            tabla_escn = (
                df_escn
                .groupby('Especialidad')
                .size()
                .reset_index(name='cantidad')
                .sort_values(by='cantidad', ascending=False)
            )  
            # =========================
            # CONTEXTO (ORIGINAL)
            # =========================
            contexto = {
                'filas': df.to_dict('records'),
                'total_es_control': total_escontrol,
                'total_es_cn': total_escn,
                'total_controles': total_controles,
                'total_consultas_nuevas': total_consultas_nuevas,
                'total_inter': total_inter,
                'resultado_es_control_menos_interconsulta': resultado,
                'especialidades': tabla.to_dict('records'),
                'total_general_control': int(tabla['cantidad'].sum()),
                'porc_escontrol_vs_controles': round((resultado/total_controles)*100,2) if total_controles else 0,
                'porc_escontrol_vs_cn': round((resultado/total_consultas_nuevas)*100,2) if total_consultas_nuevas else 0,
                'porc_escn_vs_cn': f"{round((total_escn / total_consultas_nuevas) * 100, 2)}%" if total_consultas_nuevas else "0%",
                'porc_escn_vs_controles': f"{round((total_escn / total_controles) * 100, 2)}%" if total_controles else "0%",
                'tabla_funcionarios': tabla_funcionarios.to_dict('records'),
                'tabla_escn': tabla_escn.to_dict('records'),
            }
            # =========================
            # RENDER WORD
            # =========================
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

            # Marcar que los reportes ya fueron generados
            st.session_state["reportes_generados"] = True

            st.success("✅ Reporte generado correctamente")
            # ==================================================
            # DESCARGAS (FUERA DEL BOTÓN GENERAR)
            # ==================================================

            if st.session_state.get("reportes_generados", False):

                st.divider()
                st.subheader("📥 Descargar Informes")

                colA, colB = st.columns(2)

                if "informe1" in st.session_state:
                    colA.download_button(
                        label="📥 Informe 1",
                        data=st.session_state["informe1"],
                        file_name="Informe_1.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="descarga_informe_1"
                    )

                if "informe2" in st.session_state:
                    colB.download_button(
                        label="📥 Informe 2",
                        data=st.session_state["informe2"],
                        file_name="Informe_2.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="descarga_informe_2"
                    )
         

            # =========================
            # RESULTADOS
            # =========================
            st.subheader("📊 Resultados")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Control", total_escontrol)
            c2.metric("CN", total_escn)
            c3.metric("Interconsultas", total_inter)
            c4.metric("Resultado", resultado)

            st.dataframe(tabla)

        except Exception as e:
            st.error(f"Error: {e}")

    else:
        st.warning("Debes subir los archivos base")
