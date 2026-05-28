import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile

st.set_page_config(page_title="Reporte Cuadratura", layout="centered")

st.title("📊 Generador de Reportes Médicos")

datos_file = st.file_uploader("Sube datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("Sube Lista_Espera.xlsx", type=["xlsx"])
word_file = st.file_uploader("Sube plantilla.docx", type=["docx"])


def calcular_controlycn(fila):
    actividad = str(fila.get('Actividad','')).strip()
    Ic = str(fila.get('Ic Asoc Hora','')).strip()

    rEs_control = 1 if (actividad == 'CONSULTA NUEVA' and Ic == '-') else 0
    rEs_CN = 1 if (actividad == 'CONTROL' and Ic != '-') else 0

    return pd.Series([rEs_control, rEs_CN])


def limpiar_rut_definitivo(rut):
    rut = str(rut).strip()
    if "-" in rut:
        rut = rut[:rut.find("-")]
    return rut.replace(".", "")


def marcar_interconsulta_valida(fila):
    actividad = str(fila.get('Actividad', '')).strip().upper()
    ic_asoc = str(fila.get('Ic Asoc Hora', '')).strip()
    num_ic = fila.get('Num Interconsulta', 0)

    return 1 if (actividad == 'CONSULTA NUEVA' and ic_asoc == '-' and num_ic != 0) else 0


if st.button("🚀 Generar Reporte"):

    if datos_file and lp_file and word_file:

        try:
            df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")

            doc = DocxTemplate(word_file)

            df[['Es_control', 'Es_CN']] = df.apply(calcular_controlycn, axis=1)

            total_escontrol = int(df['Es_control'].sum())
            total_escn = int(df['Es_CN'].sum())

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
            resultado_final = total_escontrol - total_inter

            df_controles = df[df['Es_control'] == 1]
            tabla = df_controles.groupby("Especialidad").size().reset_index(name="cantidad")

            # 🔥 DEBUG CLAVE
            st.write("PREVIEW FILAS WORD:")
            st.dataframe(df.head())

            contexto = {
                "filas": df.to_dict(orient="records"),
                "total_control": total_escontrol,
                "total_escn": total_escn,
                "especialidades": tabla.to_dict(orient="records"),
                "total_general_control": int(tabla['cantidad'].sum())
            }

            doc.render(contexto)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)

                st.success("✅ Reporte generado correctamente")

                with open(tmp.name, "rb") as f:
                    st.download_button(
                        "📥 Descargar Word",
                        f,
                        file_name="Reporte_Actividades.docx"
                    )

        except Exception as e:
            st.error(f"Error: {e}")

    else:
        st.warning("Debes subir los 3 archivos")
