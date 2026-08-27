import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import tempfile
from io import BytesIO
import traceback

st.set_page_config(page_title="Reporte Cuadratura", layout="wide")

st.title("📊 Generador de Reportes Médicos")

st.write("Sube los archivos base y opcionalmente las plantillas")

# =========================
# ARCHIVOS BASE
# =========================
st.markdown("## 📂 Archivos base")

datos_file = st.file_uploader("📄 Datos RCE Especialidades.xlsx", type=["xlsx"])
lp_file = st.file_uploader("📄 Lista de Espera.xlsx", type=["xlsx"])

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

st.markdown("## 📅 Fechas del informe")

col_fecha1, col_fecha2 = st.columns(2)

with col_fecha1:
    fecha_corte = st.date_input(
        "Fecha de corte"
    )

with col_fecha2:
    fecha_inf_preliminar = st.date_input(
        "Fecha informe preliminar"
    )


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
            df = pd.read_excel(
                datos_file,
                sheet_name="NOMINA CUADRATURA (REM7) SIN CO"
            )
            
            # Pandas normalmente transforma columnas duplicadas:
            # Rut -> Rut
            # Rut -> Rut.1
            
            if 'Rut' not in df.columns or 'Rut.1' not in df.columns:
                raise ValueError(
                    "No se encontraron las columnas 'Rut' y 'Rut.1'. "
                    f"Columnas encontradas: {df.columns.tolist()}"
                )
            
            # Primer Rut = paciente
            df = df.rename(columns={
                'Rut': 'Rut Paciente',
                'Rut.1': 'Rut Funcionario'
            })



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

            df['Paciente'] = (
                df['Nombre'].fillna('').astype(str) + ' ' +
                df['Apell Paterno'].fillna('').astype(str) + ' ' +
                df['Apell Materno'].fillna('').astype(str)
            ).str.replace(r'\s+', ' ', regex=True).str.strip()

            df['rut_puente'] = df['Rut Paciente'].apply(limpiar_rut_definitivo)
            



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
            resultado_escontrol = total_escontrol - total_inter

            # =========================
            # REVISION INTERCONSULTAS
            # =========================

            revision=df[
             (df['Actividad'].astype(str).str.strip().str.upper()=='CONSULTA NUEVA')&
             (df['Ic Asoc Hora'].astype(str).str.strip()=='-')
            ][
              [
                'Rut Paciente',
                'Rut Funcionario',
                'Funcionario',
                'Especialidad',
                'Actividad',
                'Ic Asoc Hora',
                'Num Interconsulta',
                'Es_control',
                'Interconsulta_Valida'
            ]

            ].copy()

            st.write("====REVISION DE INTERCONSULTAS====")
            st.write("Cantidad de casos revisador:", len(revision))
            st.write(
                "Interconsultas válidas encontradas:",
                int(revision['Interconsulta_Valida'].sum())
            )
            st.dataframe(revision)
            
            
            


            
            porc_escontrol_vs_controles = (
                resultado_escontrol / total_controles * 100
            ) if total_controles != 0 else 0
           
            porc_escontrol_vs_cn = (
                resultado_escontrol / total_consultas_nuevas * 100
            ) if total_consultas_nuevas != 0 else 0

            porc_escontrol_vs_controles = round(porc_escontrol_vs_controles, 1)
            porc_escontrol_vs_cn = round(porc_escontrol_vs_cn, 1)

            porc_escn_vs_consultas_nuevas = (
                total_es_cn / total_consultas_nuevas * 100
            ) if total_consultas_nuevas != 0 else 0

            porc_escn_vs_consultas_nuevas = round(porc_escn_vs_consultas_nuevas, 1)

            porc_escn_vs_total_controles = (
                total_es_cn / total_controles * 100
            ) if total_controles != 0 else 0

            porc_escn_vs_total_controles = round(porc_escn_vs_total_controles, 1)

           
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

            detalle_control = (
                df_controles
                .groupby(
                    [
                        'Especialidad',
                        'Rut Funcionario',
                        'Funcionario',
                        'Paciente',
                        'Fecha Atencion',
                        'Ic Asoc Hora'
                    ],
                    dropna=False
                )
                .size()
                .reset_index(name='Total')
                .sort_values(
                    [
                        'Especialidad',
                        'Funcionario',
                        'Fecha Atencion',
                        'Paciente'
                    ]
                )
            )




            detalle_escn = (
                df_escn
                .groupby(
                    [
                        'Especialidad',
                        'Rut Funcionario',
                        'Funcionario',
                        'Paciente',
                        'Fecha Atencion',
                        'Ic Asoc Hora'
                    ],
                    dropna=False
                )
                .size()
                .reset_index(name='Total')
                .sort_values(
                    [
                        'Especialidad',
                        'Funcionario',
                        'Fecha Atencion',
                        'Paciente'
                    ]
                )
            )



           
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


            fecha_corte_str = fecha_corte.strftime("%d/%m/%Y")
            fecha_inf_preliminar_str = fecha_inf_preliminar.strftime("%d/%m/%Y")
            meses = {
                1: "enero",
                2: "febrero",
                3: "marzo",
                4: "abril",
                5: "mayo",
                6: "junio",
                7: "julio",
                8: "agosto",
                9: "septiembre",
                10: "octubre",
                11: "noviembre",
                12: "diciembre"
            }
           
            mes_corte = meses[fecha_corte.month]
           # =========================
            # 📊 GENERAR EXCEL
            # =========================
            from io import BytesIO
           
            output = BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                tabla_escn.to_excel(writer, sheet_name='ES_CN_Especialidad', index=False)
                tabla_funcionarios_cn.to_excel(writer, sheet_name='ES_CN_Funcionario', index=False)
           
                detalle_escn.to_excel(
                    writer,
                    sheet_name='ES_CN_Detalle',
                    index=False
                )
           
                tabla.to_excel(writer, sheet_name='CONTROL_Especialidad', index=False)
                tabla_funcionarios.to_excel(writer, sheet_name='CONTROL_Funcionario', index=False)
           
                detalle_control.to_excel(
                    writer,
                    sheet_name='CONTROL_Detalle',
                    index=False
                )
           
            # 🔥 CLAVE: mover puntero al inicio
            output.seek(0)
           
            st.session_state["reporte_excel"] = output.read()
# =========================
# CLAVES ARCHIVO WORD
# =========================
            contexto = {
                'filas': df.to_dict('records'),
                'total_es_control': total_escontrol,
                'total_es_cn': total_escn,
                'total_controles': total_controles,
                'total_consultas_nuevas': total_consultas_nuevas,
                'total_inter': total_inter,
                'resultado_es_control_menos_interconsulta': resultado_escontrol,
                'especialidades': tabla.to_dict('records'),
                'tabla_funcionarios': tabla_funcionarios.to_dict('records'),
                'tabla_escn': tabla_escn.to_dict('records'),
                'tabla_funcionarios_cn': tabla_funcionarios_cn.to_dict('records'),
                'porc_escontrol_vs_controles': porc_escontrol_vs_controles,
                'porc_escontrol_vs_cn': porc_escontrol_vs_cn,
                'porc_escn_vs_consultas_nuevas': porc_escn_vs_consultas_nuevas,
                'porc_escn_vs_controles': porc_escn_vs_total_controles,
                'fecha_corte': fecha_corte_str,
                'fecha_inf_preliminar': fecha_inf_preliminar_str,
                'mes_corte': mes_corte,
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

colA, colB, colC = st.columns(3)

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

with colC:
    if st.session_state.get("reporte_excel"):
   
        st.download_button(
            "📥 Descargar Excel consolidado",
            data=st.session_state["reporte_excel"],
            file_name="Reporte_Consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
