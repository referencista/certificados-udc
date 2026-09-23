import io
import os
import re
import datetime
import requests
import pandas as pd
import streamlit as st

# Intentar importar librerías opcionales/externas
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSHEETS = True
except ImportError:
    HAS_GSHEETS = False

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==========================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS (COLORS UCUENCA)
# ==========================================
st.set_page_config(
    page_title="Gestión de Certificados UCuenca",
    page_icon="🎓",
    layout="wide"
)

# Inyección de CSS para personalizar los colores de la interfaz
st.markdown("""
    <style>
    /* Color de fondo general de la app */
    .stApp {
        background-color: #F8F9FA;
    }
    
    /* Encabezados principales con Azul Institucional UCuenca */
    h1, h2, h3 {
        color: #002060 !important;
        font-weight: 700 !important;
    }
    
    /* Estilo de los Botones Principales */
    div.stButton > button {
        background-color: #002060 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }
    
    div.stButton > button:hover {
        background-color: #001030 !important;
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.15) !important;
    }

    /* Estilo para las Pestañas (Tabs) */
    button[data-baseweb="tab"] {
        font-size: 16px !important;
        font-weight: 600 !important;
        color: #002060 !important;
    }
    
    /* Estilos para las tarjetas de KPIs */
    [data-testid="stMetricValue"] {
        font-size: 30px !important;
        color: #002060 !important;
        font-weight: bold !important;
    }

    [data-testid="stMetricLabel"] {
        font-size: 14px !important;
        color: #555555 !important;
        font-weight: 600 !important;
    }

    div[data-testid="metric-container"] {
        background-color: #FFFFFF !important;
        padding: 18px !important;
        border-radius: 10px !important;
        border-left: 5px solid #002060 !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06) !important;
    }
    </style>
""", unsafe_allow_html=True)

COLUMNAS_GSHEETS = ["Fecha y Hora", "Referencista", "Estudiante", "Facultad", "Carrera", "Handle"]
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ==========================================
# CONEXIÓN SECURIZADA A GOOGLE SHEETS
# ==========================================
@st.cache_resource
def conectar_google_sheets():
    if not HAS_GSHEETS:
        return None
    try:
        # 1. Intentar por secrets de Streamlit
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=SCOPES
            )
            client = gspread.authorize(creds)
            sheet_name = st.secrets.get("GSHEET_NAME", "Certificados_UCuenca")
            return client.open(sheet_name).sheet1

        # 2. Intentar por archivo local solo si existe físicamente
        elif os.path.exists("credentials.json"):
            creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
            client = gspread.authorize(creds)
            sheet_name = st.secrets.get("GSHEET_NAME", "Certificados_UCuenca")
            return client.open(sheet_name).sheet1

    except Exception:
        pass
    return None

def guardar_en_google_sheets(sheet, datos_dict):
    if sheet is None:
        return False
    try:
        registros_existentes = sheet.get_all_values()
        if len(registros_existentes) == 0:
            sheet.append_row(COLUMNAS_GSHEETS)
        
        fila = [datos_dict.get(col, "") for col in COLUMNAS_GSHEETS]
        sheet.append_row(fila)
        return True
    except Exception as e:
        st.error(f"Error al guardar registro: {e}")
        return False

def cargar_datos_reporte(sheet):
    if sheet is None:
        return pd.DataFrame(columns=COLUMNAS_GSHEETS)
    try:
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        for col in COLUMNAS_GSHEETS:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_GSHEETS)

# ==========================================
# CONSULTA A DSPACE
# ==========================================
def consultar_dspace(url_or_handle):
    match = re.search(r'handle/(\d+/\d+)', url_or_handle)
    handle_id = match.group(1) if match else url_or_handle.strip()

    datos = {
        "Estudiante": "",
        "Facultad": "",
        "Carrera": "",
        "Handle": handle_id,
        "Titulo": ""
    }

    try:
        url_target = url_or_handle if url_or_handle.startswith("http") else f"https://dspace.ucuenca.edu.ec/handle/{handle_id}"
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url_target, headers=headers, timeout=10)
        
        if response.status_code == 200:
            html = response.text
            autores = re.findall(r'<meta name="DC.creator" content="([^"]+)"', html)
            if not autores:
                autores = re.findall(r'<meta name="citation_author" content="([^"]+)"', html)
            if autores:
                datos["Estudiante"] = ", ".join(autores)
                
            titulos = re.findall(r'<meta name="DC.title" content="([^"]+)"', html)
            if titulos:
                datos["Titulo"] = titulos[0]

            facultades = re.findall(r'<meta name="DC.publisher" content="([^"]+)"', html)
            if facultades:
                datos["Facultad"] = facultades[0]

    except Exception as e:
        st.warning(f"No se pudieron extraer metadatos automáticamente ({e}). Completa los campos a mano.")

    return datos

# ==========================================
# GENERADOR WORD
# ==========================================
def generar_word_doc(estudiante, facultad, carrera, referencista, handle):
    if not HAS_DOCX:
        st.error("La librería python-docx no está instalada.")
        return None

    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("UNIVERSIDAD DE CUENCA\nCENTRO DE INFORMACIÓN Y DOCUMENTACIÓN")
    run_title.bold = True
    run_title.font.size = Pt(14)
    run_title.font.color.rgb = RGBColor(0, 32, 96)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("CERTIFICADO DE NO ADEUDAR TRABAJO DE TITULACIÓN / DEPÓSITO DIGITAL")
    run_sub.bold = True
    run_sub.font.size = Pt(12)

    doc.add_paragraph().paragraph_format.space_after = Pt(24)

    fecha_actual = datetime.datetime.now().strftime("%d de %B de %Y")
    
    p_body = doc.add_paragraph()
    p_body.paragraph_format.line_spacing = 1.15
    p_body.paragraph_format.space_after = Pt(12)
    
    p_body.add_run("Por medio del presente se hace constar que el/la estudiante ")
    p_body.add_run(f"{estudiante}").bold = True
    p_body.add_run(", perteneciente a la Facultad de ")
    p_body.add_run(f"{facultad}").bold = True
    p_body.add_run(", Carrera de ")
    p_body.add_run(f"{carrera}").bold = True
    p_body.add_run(", ha cumplido satisfactoriamente con la entrega e ingreso en el Repositorio Digital Institucional (DSpace) de su trabajo de titulación.\n\n")
    p_body.add_run("Handle / Identificador de Registro: ").bold = True
    p_body.add_run(f"{handle}\n")

    doc.add_paragraph().paragraph_format.space_after = Pt(36)

    p_fecha = doc.add_paragraph()
    p_fecha.add_run(f"Cuenca, {fecha_actual}")

    doc.add_paragraph().paragraph_format.space_after = Pt(48)

    p_firma = doc.add_paragraph()
    p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_firma.add_run(f"___________________________________\n{referencista}\n").bold = True
    p_firma.add_run("Referencista / Centro de Información y Documentación\nUniversidad de Cuenca")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def resetear_estado_busqueda():
    st.session_state.pop("datos_cargados", None)

# ==========================================
# APP PRINCIPAL
# ==========================================
def main():
    st.title("🎓 Sistema de Gestión de Certificados - UCuenca")

    sheet = conectar_google_sheets()

    tab1, tab2 = st.tabs(["📄 Generar Certificado", "📊 Dashboard y Reportes"])

    with tab1:
        st.header("📄 Generador de Certificados DSpace")
        st.caption("Ingrese la URL o Handle para obtener los metadatos institucionales.")

        url_dspace = st.text_input(
            "URL o Handle de DSpace:",
            key="input_url",
            placeholder="Ej: https://dspace.ucuenca.edu.ec/handle/123456789/12345",
            on_change=resetear_estado_busqueda
        )

        btn_procesar = st.button("🔍 Consultar DSpace")

        if btn_procesar:
            if not url_dspace.strip():
                st.error("Por favor, ingrese una URL o Handle válido.")
            else:
                with st.spinner("Consultando metadatos en DSpace..."):
                    datos = consultar_dspace(url_dspace)
                    st.session_state["datos_cargados"] = datos
                    st.success("Metadatos procesados correctamente.")

        if "datos_cargados" in st.session_state:
            datos = st.session_state["datos_cargados"]
            
            st.markdown("---")
            st.subheader("📝 Confirmar y Editar Información")

            col1, col2 = st.columns(2)
            with col1:
                estudiante = st.text_input("Estudiante / Autor(es):", value=datos.get("Estudiante", ""))
                facultad = st.text_input("Facultad:", value=datos.get("Facultad", ""))
            
            with col2:
                carrera = st.text_input("Carrera:", value=datos.get("Carrera", ""))
                referencista = st.selectbox(
                    "Referencista Emisor:",
                    ["Referencista 1", "Referencista 2", "Referencista 3", "Librarian Admin"]
                )

            handle_actual = datos.get("Handle", "")
            st.caption(f"**Handle:** `{handle_actual}`")

            if st.button("🛠️ Generar Certificado y Registrar"):
                if not estudiante or not facultad or not carrera:
                    st.warning("Por favor complete todos los campos requeridos.")
                else:
                    docx_buffer = generar_word_doc(
                        estudiante=estudiante,
                        facultad=facultad,
                        carrera=carrera,
                        referencista=referencista,
                        handle=handle_actual
                    )

                    registro = {
                        "Fecha y Hora": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Referencista": referencista,
                        "Estudiante": estudiante,
                        "Facultad": facultad,
                        "Carrera": carrera,
                        "Handle": handle_actual
                    }
                    
                    guardado_ok = guardar_en_google_sheets(sheet, registro)

                    if guardado_ok:
                        st.success("✅ Registro almacenado en Google Sheets.")
                    else:
                        st.info("ℹ️ Certificado generado exitosamente.")

                    if docx_buffer:
                        st.download_button(
                            label="📥 Descargar Certificado (.docx)",
                            data=docx_buffer,
                            file_name=f"Certificado_{estudiante.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

    with tab2:
        st.header("📊 Dashboard de Certificados Emitidos")

        df_registros = cargar_datos_reporte(sheet)

        if df_registros.empty:
            st.info("No hay registros almacenados actualmente en la base de datos.")
        else:
            df_registros['Fecha_DT'] = pd.to_datetime(df_registros['Fecha y Hora'], errors='coerce')
            df_registros['Año'] = df_registros['Fecha_DT'].dt.year.fillna(0).astype(int)
            df_registros['Mes'] = df_registros['Fecha_DT'].dt.strftime('%B')
            df_registros['Fecha_Corta'] = df_registros['Fecha_DT'].dt.date

            st.subheader("🔍 Filtros de Búsqueda")
            col_f1, col_f2, col_f3 = st.columns(3)

            with col_f1:
                anios_validos = [a for a in df_registros['Año'].unique() if a != 0]
                anios_disponibles = ["Todos"] + sorted(anios_validos, reverse=True)
                anio_sel = st.selectbox("Año:", anios_disponibles)

            with col_f2:
                meses_disponibles = ["Todos"] + [m for m in df_registros['Mes'].unique() if pd.notna(m)]
                mes_sel = st.selectbox("Mes:", meses_disponibles)

            with col_f3:
                refs_disponibles = ["Todos"] + [r for r in df_registros['Referencista'].unique() if r]
                ref_filtro = st.selectbox("Referencista:", refs_disponibles)

            df_filtrado = df_registros.copy()

            if anio_sel != "Todos":
                df_filtrado = df_filtrado[df_filtrado['Año'] == anio_sel]

            if mes_sel != "Todos":
                df_filtrado = df_filtrado[df_filtrado['Mes'] == mes_sel]

            if ref_filtro != "Todos":
                df_filtrado = df_filtrado[df_filtrado['Referencista'] == ref_filtro]

            tot_certificados = len(df_filtrado)

            if not df_filtrado.empty and 'Referencista' in df_filtrado.columns:
                top_ref_series = df_filtrado['Referencista'].mode()
                referencista_top = top_ref_series[0] if not top_ref_series.empty else "N/A"
            else:
                referencista_top = "N/A"

            if not df_filtrado.empty and 'Fecha_Corta' in df_filtrado.columns:
                dias_activos = df_filtrado['Fecha_Corta'].nunique()
                promedio_diario = round(tot_certificados / dias_activos, 1) if dias_activos > 0 else 0
            else:
                promedio_diario = 0

            st.markdown("---")
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Total Certificados", tot_certificados)
            kpi2.metric("Referencista Top", referencista_top)
            kpi3.metric("Promedio Diario", f"{promedio_diario} / día")

            st.markdown("---")

            st.subheader("📋 Registros Detallados")
            columnas_visibles = ["Fecha y Hora", "Referencista", "Estudiante", "Facultad", "Carrera", "Handle"]
            
            st.dataframe(
                df_filtrado[columnas_visibles],
                use_container_width=True,
                hide_index=True
            )

            csv_data = df_filtrado[columnas_visibles].to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Exportar Reporte Filtrado (CSV)",
                data=csv_data,
                file_name=f"reporte_certificados_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )

if __name__ == "__main__":
    main()
