import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile
from io import BytesIO

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")

st.write("Sube los archivos base y opcionalmente las plantillas")

# =========================
# ESTADO STREAMLIT (SOLO DESCARGA)
# =========================
if "generado" not in st.session_state:
    st.session_state["generado"] = False

if "informe1" not in st.session_state:
    st.session_state["informe1"] = None

if "informe2" not in st.session_state:
    st.session_state["informe2"] = None

if "excel" not in st.session_state:
    st.session_state["excel"] = None

# =========================
# ARCHIVOS BASE
# =========================
st.markdown("## 📂 Archivos base")

datos_file = st.file_uploader("📄 Subir datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("📄 Subir Lista_Espera.xlsx", type=["xlsx"])

st.divider()

# =========================
# FUNCIONES (ORIGINALES - NO CAMBIADAS)
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
    word_file = st.file_uploader("Plantilla 1 (.docx)", type=["docx"], key="w1")

with col2:
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
            # ES_CONTROL / ES_CN
            # =========================
            df[['Es_control', 'Es_CN']] = df.apply(calcular_controlycn, axis=1)

            total_escontrol = int(df['Es_control'].sum())
            total_escn = int(df['Es_CN'].sum())

            # =========================
            # FUNCIONARIO
            # =========================
            df['Funcionario'] = (
                df['Nombres'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Pat'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Mat'].fillna('').astype(str).str.strip()
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

            # =========================
            # INTERCONSULTA
            # =========================
            df['Interconsulta_Valida'] = df.apply(marcar_interconsulta_valida, axis=1)

            total_inter = int(df['Interconsulta_Valida'].sum())
            
            total_es_control_real = total_escontrol - total_inter
        

            # =========================
            # 🔴 ERROR (ESTO ES LO QUE FALTABA EN ORDEN)
            # =========================
            df['Error_escontrol_menos_Inter'] = (df['Es_control'] - df['Interconsulta_Valida'])
            df['Error_escontrol_menos_Inter'] = (df['Error_escontrol_menos_Inter'] > 0).astype(int)

            # =========================
            # FILTROS (DESPUÉS DEL ERROR)
            # =========================
            df_controles = df[df['Error_escontrol_menos_Inter'] == 1]
            df_escn = df[df['Es_CN'] == 1]

            # =========================
            # TABLAS
            # =========================
            tabla = (
                df[(df['Es_control'] == 1) & (df['Interconsulta_Valida'] == 0)]
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
            
            tabla_funcionarios_escn = (
                df_escn
                .groupby(['Especialidad', 'Funcionario'])
                .size()
                .reset_index(name='total')
                .sort_values(by='total', ascending=False)
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
                'resultado_es_control_menos_interconsulta': total_es_control_real,
                'especialidades': tabla.to_dict('records'),
                'total_general_control': int(tabla['cantidad'].sum()),
                'porc_escontrol_vs_controles': round((total_es_control_real / total_controles) * 100, 2) if total_controles else 0,
                'porc_escontrol_vs_cn': round((total_es_control_real / total_consultas_nuevas) * 100, 2) if total_consultas_nuevas else 0,
                'porc_escn_vs_cn': f"{round((total_escn / total_consultas_nuevas) * 100, 2)}%" if total_consultas_nuevas else "0%",
                'porc_escn_vs_controles': f"{round((total_escn / total_controles) * 100, 2)}%" if total_controles else "0%",
                'tabla_funcionarios': tabla_funcionarios.to_dict('records'),
                'tabla_escn': tabla_escn.to_dict('records'),
                'tabla_funcionarios_escn': tabla_funcionarios_escn.to_dict('records'),
            }



            resumen = pd.DataFrame({
                "Indicador": [
                    "Total Controles",
                    "Total Consultas Nuevas",
                    "Total ES_CONTROL",
                    "Total ES_CN",
                    "Total Interconsultas Válidas",
                    "Total ES_CONTROL Real"
                ],
                "Valor": [
                    total_controles,
                    total_consultas_nuevas,
                    total_escontrol,
                    total_escn,
                    total_inter,
                    total_es_control_real
                ]
            })
            # =========================
            # TABLAS PARA EXCEL
            # =========================
            tabla_control_excel = tabla.rename(columns={
                "cantidad": "total_es_control_real"
            })
            
            tabla_funcionarios_control_excel = tabla_funcionarios.rename(columns={
                "total": "total_es_control_real"
            })
            
            tabla_escn_excel = tabla_escn.rename(columns={
                "cantidad": "total_es_cn"
            })
            
            tabla_funcionarios_escn_excel = tabla_funcionarios_escn.rename(columns={
                "total": "total_es_cn"
            })
            # =========================
            # RENDER WORD (SOLO CAMBIO AQUÍ)
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

            # =========================
            # GENERAR EXCEL
            # =========================
            excel_buffer = BytesIO()

            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:

                resumen.to_excel(
                    writer,
                    sheet_name="Resumen",
                    index=False
                )
            
                tabla_control_excel.to_excel(
                    writer,
                    sheet_name="Control_Especialidad",
                    index=False
                )
            
                tabla_funcionarios_control_excel.to_excel(
                    writer,
                    sheet_name="Control_Funcionario",
                    index=False
                )
            
                tabla_escn_excel.to_excel(
                    writer,
                    sheet_name="CN_Especialidad",
                    index=False
                )
            
                tabla_funcionarios_escn_excel.to_excel(
                    writer,
                    sheet_name="CN_Funcionario",
                    index=False
                )

            excel_buffer.seek(0)

            st.session_state["excel"] = excel_buffer.getvalue()

            
            st.session_state["generado"] = True

            st.success("✅ Reporte generado correctamente")

        except Exception as e:
            st.error(f"Error: {e}")

    else:
        st.warning("Debes subir los archivos base")


# =========================
# DESCARGAS (CORREGIDO)
# =========================
if st.session_state.get("generado", False):

    st.divider()
    st.subheader("📥 Descargar Informes")

    colA, colB, colC = st.columns(3)

    if st.session_state.get("informe1"):
        colA.download_button(
            "📥 Informe 1",
            data=st.session_state["informe1"],
            file_name="Informe_1.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl1"
        )

    if st.session_state.get("informe2"):
        colB.download_button(
            "📥 Informe 2",
            data=st.session_state["informe2"],
            file_name="Informe_2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl2"
        )
    if st.session_state.get("excel"):
        colC.download_button(
            "📊 Reporte Excel",
            data=st.session_state["excel"],
            file_name="Reporte_Cuadratura.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_excel"
        )
