import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile

st.set_page_config(page_title="Reporte Cuadratura", layout="centered")

st.title("📊 Generador de Reportes Médicos")

st.write("Sube los archivos para generar el reporte automáticamente")

# 📂 Subida de archivos
datos_file = st.file_uploader("Sube datos.xlsx", type=["xlsx"])
lp_file = st.file_uploader("Sube Lista_Espera.xlsx", type=["xlsx"])
word_file = st.file_uploader("Sube plantilla.docx", type=["docx"])


# ---------------- FUNCIONES ----------------

def calcular_controlycn(fila):
    actividad = str(fila.get('Actividad', '')).strip()
    ic = str(fila.get('Ic Asoc Hora', '')).strip()

    rEs_control = 1 if (actividad == 'CONSULTA NUEVA' and ic == '-') else 0
    rEs_cn = 1 if (actividad == 'CONTROL' and ic != '-') else 0

    return pd.Series([rEs_control, rEs_cn])


def limpiar_rut_definitivo(rut):
    rut = str(rut).strip()
    if "-" in rut:
        rut = rut.split("-")[0]
    return rut.replace(".", "")


def marcar_interconsulta_valida(fila):
    actividad = str(fila.get('Actividad', '')).strip().upper()
    ic_asoc = str(fila.get('Ic Asoc Hora', '')).strip()
    num_ic = fila.get('Num Interconsulta', 0)

    if actividad == 'CONSULTA NUEVA' and ic_asoc == '-' and num_ic != 0:
        return 1
    return 0


# ---------------- EJECUCIÓN ----------------

if st.button("🚀 Generar Reporte"):

    if datos_file and lp_file and word_file:

        try:
            # ---------------- LEER ARCHIVOS ----------------
            df = pd.read_excel(datos_file, sheet_name="NOMINA CUADRATURA (REM7) SIN CO")
            lp = pd.read_excel(lp_file, sheet_name="Nomina Médico")

            doc = DocxTemplate(word_file)

            # ---------------- CÁLCULOS ----------------
            df[['Es_control', 'Es_cn']] = df.apply(calcular_controlycn, axis=1)

            total_control = int(df['Es_control'].sum())
            total_cn = int(df['Es_cn'].sum())

            df['rut_puente'] = df['Rut'].apply(limpiar_rut_definitivo)
            lp['rut_puente'] = lp['Rut'].apply(limpiar_rut_definitivo)

            df['esp_puente'] = df['Especialidad'].str.strip().str.upper()
            lp['esp_puente'] = lp['Especialidad Destino'].str.strip().str.upper()

            df = pd.merge(
                df,
                lp[['rut_puente', 'esp_puente', 'Num Interconsulta']],
                on=['rut_puente', 'esp_puente'],
                how='left'
            )

            df['Num Interconsulta'] = df['Num Interconsulta'].fillna(0)
            df['Interconsulta_Valida'] = df.apply(marcar_interconsulta_valida, axis=1)

            total_inter = int(df['Interconsulta_Valida'].sum())
            resultado_final = total_control - total_inter

            # ---------------- TABLA RESUMEN ----------------
            df_controles = df[df['Es_control'] == 1]
            tabla = df_controles.groupby("Especialidad").size().reset_index(name="cantidad")

            # ---------------- CONTEXTO WORD ----------------
            filas = df.to_dict(orient='records')
            especialidades = tabla.to_dict(orient='records')

            contexto = {
                "filas": filas,
                "total_control": total_control,
                "total_cn": total_cn,
                "especialidades": especialidades,
                "total_general_control": int(tabla["cantidad"].sum())
            }

            # 🔥 DEBUG (ordenado)
            st.subheader("🔍 Debug contexto Word")
            st.json(contexto)

            # ---------------- RENDER WORD ----------------
            doc.render(contexto)

            # ---------------- GUARDAR ARCHIVO ----------------
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)

                st.success("✅ Reporte generado correctamente")

                with open(tmp.name, "rb") as f:
                    st.download_button(
                        "📥 Descargar Word",
                        f,
                        file_name="Reporte_Actividades.docx"
                    )

            # ---------------- RESULTADOS ----------------
            st.subheader("📊 Resultados")
            st.write("Control:", total_control)
            st.write("CN:", total_cn)
            st.write("Interconsultas:", total_inter)
            st.write("Resultado final:", resultado_final)

            st.subheader("📌 Resumen por especialidad")
            st.dataframe(tabla, use_container_width=True)

        except Exception as e:
            st.error(f"Error: {e}")

    else:
        st.warning("Debes subir los 3 archivos")
