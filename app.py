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

    if actividad == 'CONSULTA NUEVA' and ic_asoc == '-' and num_ic != 0:
        return 1
    return 0


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

            # =========================================================
            # 🔥 FIX CLAVE: CRUCE DE RUT + NUM INTERCONSULTA
            # =========================================================
            df['Rut_Limpio'] = df['Rut'].apply(limpiar_rut_definitivo)
            lp['Rut_Limpio'] = lp['Rut'].apply(limpiar_rut_definitivo)

            lp['Num Interconsulta'] = pd.to_numeric(
                lp['Num Interconsulta'],
                errors='coerce'
            ).fillna(0)

            df = df.merge(
                lp[['Rut_Limpio', 'Num Interconsulta']],
                on='Rut_Limpio',
                how='left'
            )

            df['Num Interconsulta'] = df['Num Interconsulta'].fillna(0)

            # =========================
            # TOTALES
            # =========================
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
                df['Nombres'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Pat'].fillna('').astype(str).str.strip() + ' ' +
                df['Apellido Mat'].fillna('').astype(str).str.strip()
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

            # =========================
            # INTERCONSULTA
            # =========================
            df['Interconsulta_Valida'] = df.apply(marcar_interconsulta_valida, axis=1)

            st.subheader("🔍 Interconsultas válidas detectadas")

            inter_validas = df[df['Interconsulta_Valida'] == 1]

            st.write(f"Total detectadas: {len(inter_validas)}")

            st.dataframe(
                inter_validas[
                    [
                        'Actividad',
                        'Ic Asoc Hora',
                        'Num Interconsulta',
                        'Especialidad',
                        'Nombres',
                        'Apellido Pat',
                        'Apellido Mat'
                    ]
                ]
            )

            total_inter = int(df['Interconsulta_Valida'].sum())

            total_es_control_real = total_escontrol - total_inter

            # =========================
            # ERROR ES CONTROL
            # =========================
            df['Error_escontrol_menos_Inter'] = (df['Es_control'] - df['Interconsulta_Valida'])
            df['Error_escontrol_menos_Inter'] = (df['Error_escontrol_menos_Inter'] > 0).astype(int)

            # =========================
            # FILTROS
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
            # CONTEXTO
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
            }

            # =========================
            # EXCEL
            # =========================
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

            excel_buffer = BytesIO()

            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                resumen.to_excel(writer, sheet_name="Resumen", index=False)

            excel_buffer.seek(0)
            st.session_state["excel"] = excel_buffer.getvalue()

            st.session_state["generado"] = True
            st.success("✅ Reporte generado correctamente")

        except Exception as e:
            st.error(f"Error: {e}")

    else:
        st.warning("Debes subir los archivos base")


# =========================
# DESCARGAS
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
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    if st.session_state.get("informe2"):
        colB.download_button(
            "📥 Informe 2",
            data=st.session_state["informe2"],
            file_name="Informe_2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    if st.session_state.get("excel"):
        colC.download_button(
            "📊 Reporte Excel",
            data=st.session_state["excel"],
            file_name="Reporte_Cuadratura.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
