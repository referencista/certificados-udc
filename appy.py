import io
import re
from datetime import datetime
from bs4 import BeautifulSoup
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
import requests
import streamlit as st

# Configuración de la página web
st.set_page_config(
    page_title="Generador de Certificados - UCuenca", 
    page_icon="📜"
)

st.title("📜 Generador Automático de Certificado de No Adeudar")
st.subheader("Centro de Documentación Regional 'Juan Bautista Vázquez'")

# ------------------------------------------------------------------
# LISTA OFICIAL DE REFERENCISTAS
# ------------------------------------------------------------------
LISTA_REFERENCISTAS = [
    {"nombre": "DORIS PATRICIA TENESACA CARDENAS", "cargo": "Bibliotecario"},
    {"nombre": "ERIKA ELIZABETH IDROVO SALAZAR", "cargo": "Bibliotecario"},
    {"nombre": "ERIKA SOFIA PEÑAFIEL VAZQUEZ", "cargo": "Bibliotecario"},
    {"nombre": "FRANCISCO TEODORO ASTUDILLO SAQUINAULA", "cargo": "Bibliotecario"},
    {"nombre": "JENNY EULALIA PEREZ MEJIA", "cargo": "Bibliotecario"},
    {"nombre": "JHOANNA NOEMI MOGOLLON GUZMAN", "cargo": "Bibliotecario"},
    {"nombre": "PAOLA DEL ROCIO AMAYA ARCE", "cargo": "Bibliotecario 2"},
    {"nombre": "PATRICIA MARIBEL DUCHI PESANTEZ", "cargo": "Bibliotecario"},
    {"nombre": "WILMAN GONZALO TANDAZO GUEVARA", "cargo": "Bibliotecario"},
]

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

# Inicialización de la sesión para evitar reinicios en Edge
if "cert_buffer" not in st.session_state:
    st.session_state["cert_buffer"] = None
if "cert_filename" not in st.session_state:
    st.session_state["cert_filename"] = ""
if "cert_info" not in st.session_state:
    st.session_state["cert_info"] = None

# ------------------------------------------------------------------
# EXTRACCIÓN AUTOMÁTICA MULTI-MÉTODO DESDE DSPACE
# ------------------------------------------------------------------
def extraer_metadatos_dspace(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    
    # Extraer Handle o UUID de la URL
    match_handle = re.search(r'(\d+/\d+)', url)
    match_uuid = re.search(r'/items/([a-f0-9\-]+)', url)

    autor, facultad, carrera, handle_final = "", "", "", url

    # MÉTODO 1: Búsqueda vía REST API de DSpace
    try:
        if match_uuid:
            api_url = f"https://dspace.ucuenca.edu.ec/server/api/core/items/{match_uuid.group(1)}"
        elif match_handle:
            api_url = f"https://dspace.ucuenca.edu.ec/server/api/discover/search/objects?query=handle:{match_handle.group(1)}"
        else:
            api_url = None

        if api_url:
            resp = requests.get(api_url, headers=headers, timeout=12)
            if resp.status_code == 200:
                data = resp.json()
                if "_embedded" in data and "searchResult" in data["_embedded"]:
                    objects = data["_embedded"]["searchResult"]["_embedded"]["objects"]
                    item_data = objects[0]["_embedded"]["indexableObject"] if objects else {}
                else:
                    item_data = data

                metadata = item_data.get("metadata", {})
                autor = metadata.get("dc.contributor.author", [{}])[0].get("value", "")
                facultad = metadata.get("thesis.degree.grantor", [{}])[0].get("value", "")
                carrera = metadata.get("thesis.degree.discipline", [{}])[0].get("value", "")
                handle_final = metadata.get("dc.identifier.uri", [{}])[0].get("value", url)
    except Exception:
        pass # Si falla la API, pasa de inmediato al Método 2 (HTML Directo)

    # MÉTODO 2: Raspado HTML (HTML Meta-tags) como respaldo
    if not autor or not facultad:
        try:
            resp_html = requests.get(url, headers=headers, timeout=12)
            if resp_html.status_code == 200:
                soup = BeautifulSoup(resp_html.text, 'html.parser')
                
                # Buscar autor en metatags
                meta_author = soup.find('meta', {'name': re.compile(r'DC\.creator|citation_author', re.I)})
                if meta_author and meta_author.get('content'):
                    autor = meta_author['content']

                # Buscar facultad/carrera en metatags o contenido
                meta_publisher = soup.find('meta', {'name': re.compile(r'DC\.publisher|citation_publisher', re.I)})
                if meta_publisher and meta_publisher.get('content'):
                    text_pub = meta_publisher['content']
                    if "Facultad" in text_pub:
                        facultad = text_pub

                meta_uri = soup.find('meta', {'name': re.compile(r'DC\.identifier|citation_abstract_html_url', re.I)})
                if meta_uri and meta_uri.get('content'):
                    handle_final = meta_uri['content']
        except Exception as e:
            st.error(f"Error al conectar con DSpace: {str(e)}")

    if not autor:
        return None, "No se pudieron extraer los metadatos de DSpace. Revisa que la URL sea pública y correcta."

    return {
        "autor": autor.upper().strip(),
        "facultad": facultad if facultad else "Facultad de Ciencias Químicas",
        "carrera": carrera if carrera else "Carrera de Graduación",
        "handle": handle_final
    }, None

# ------------------------------------------------------------------
# GENERADOR DEL DOCUMENTO WORD (.DOCX)
# ------------------------------------------------------------------
def crear_documento_word(datos):
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # 1. ENCABEZADO
    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False

    cell_left = table_header.rows[0].cells[0]
    cell_right = table_header.rows[0].cells[1]
    cell_left.width = Inches(3.2)
    cell_right.width = Inches(3.3)

    p_logo = cell_left.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_logo = p_logo.add_run("UCUENCA")
    run_logo.font.name = 'Arial'; run_logo.font.size = Pt(28); run_logo.font.bold = True
    run_logo.font.color.rgb = RGBColor(15, 43, 91)

    p_hdr = cell_right.paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = p_hdr.add_run("FORMATO DE NO ADEUDAR MATERIAL BIBLIOGRÁFICO A LA BIBLIOTECA\n")
    r1.font.bold = True; r1.font.size = Pt(8.5); r1.font.name = 'Arial'
    r2 = p_hdr.add_run("UC-CDRJVB-FOR-020\n")
    r2.font.bold = True; r2.font.size = Pt(8.5); r2.font.name = 'Arial'
    r3 = p_hdr.add_run("Página 1 de 1")
    r3.font.size = Pt(8.5); r3.font.name = 'Arial'

    doc.add_paragraph()

    # 2. TÍTULO
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(24)
    p_titulo.paragraph_format.space_after = Pt(24)
    run_titulo = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
    run_titulo.font.name = 'Arial'; run_titulo.font.size = Pt(13); run_titulo.font.bold = True

    # 3. CUERPO
    p_cuerpo = doc.add_paragraph()
    p_cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cuerpo.paragraph_format.line_spacing = 1.15
    p_cuerpo.paragraph_format.space_after = Pt(24)

    prefix_carrera = "de la Carrera de" if datos["nivel"] == "Pregrado" else "del Programa de"

    p_cuerpo.add_run('El Centro de Documentación Regional "Juan Bautista Vázquez" certifica que ').font.name = 'Arial'
    p_cuerpo.runs[0].font.size = Pt(11)

    r_nombre = p_cuerpo.add_run(f'{datos["autor"]}')
    r_nombre.font.name = 'Arial'; r_nombre.font.size = Pt(11); r_nombre.font.bold = True

    r = p_cuerpo.add_run(', portador de la cédula de ciudadanía No. ')
    r.font.name = 'Arial'; r.font.size = Pt(11)

    r_cedula = p_cuerpo.add_run(f'{datos["cedula"]}')
    r_cedula.font.name = 'Arial'; r_cedula.font.size = Pt(11); r_cedula.font.bold = True

    r_text = f', estudiante de la {datos["facultad"]} {prefix_carrera} {datos["carrera"]}, no adeuda ningún bien, ni material bibliográfico en esta dependencia.'
    r = p_cuerpo.add_run(r_text)
    r.font.name = 'Arial'; r.font.size = Pt(11)

    # 4. FECHA
    p_fecha = doc.add_paragraph()
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_fecha.paragraph_format.space_after = Pt(28)
    hoy = datetime.now()
    fecha_texto = f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"
    r_fecha = p_fecha.add_run(fecha_texto)
    r_fecha.font.name = 'Arial'; r_fecha.font.size = Pt(11)

    # 5. FIRMA
    p_atentamente = doc.add_paragraph()
    p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_atentamente.paragraph_format.space_after = Pt(40)
    p_atentamente.add_run("Atentamente,").font.name = 'Arial'

    p_linea = doc.add_paragraph()
    p_linea.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_linea.paragraph_format.space_after = Pt(4)
    p_linea.add_run("________________________________________").font.name = 'Arial'

    p_firma = doc.add_paragraph()
    p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_firma.paragraph_format.line_spacing = 1.1

    r_nom = p_firma.add_run(f'{datos["ref_nombre"]}\n')
    r_nom.font.name = 'Arial'; r_nom.font.size = Pt(11); r_nom.font.bold = True

    r_cargo = p_firma.add_run(f'{datos["ref_cargo"]}\n')
    r_cargo.font.name = 'Arial'; r_cargo.font.size = Pt(10)

    r_cdr = p_firma.add_run('CDR "Juan Bautista Vázquez"')
    r_cdr.font.name = 'Arial'; r_cdr.font.size = Pt(10)

    for _ in range(3):
        doc.add_paragraph()

    # 6. PIE DE PÁGINA
    table_footer = doc.add_table(rows=1, cols=2)
    table_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_footer.autofit = False

    cell_f_left = table_footer.rows[0].cells[0]
    cell_f_right = table_footer.rows[0].cells[1]
    cell_f_left.width = Inches(5.0)
    cell_f_right.width = Inches(1.5)

    p_link = cell_f_left.paragraphs[0]
    p_link.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_link.add_run("Link: ").font.name = 'Arial'
    
    r_l2 = p_link.add_run(datos["handle"])
    r_l2.font.name = 'Arial'; r_l2.font.size = Pt(9.5); r_l2.font.underline = True
    r_l2.font.color.rgb = RGBColor(0, 51, 153)

    p_ver = cell_f_right.paragraphs[0]
    p_ver.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_v = p_ver.add_run("Version: 2.0")
    r_v.font.name = 'Arial'; r_v.font.size = Pt(8.5)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------------
# INTERFAZ DE USUARIO STREAMLIT
# ------------------------------------------------------------------
st.markdown("---")

col1, col2 = st.columns([2, 1])

with col1:
    url_dspace = st.text_input(
        "URL del Ítem en DSpace:", 
        placeholder="https://dspace.ucuenca.edu.ec/handle/123456789/..."
    )

with col2:
    cedula = st.text_input("Número de Cédula:", placeholder="1401064256")

col_lvl, col_ref = st.columns(2)

with col_lvl:
    nivel_academico = st.selectbox("Nivel Académico:", ["Pregrado", "Maestría / Posgrado"])

with col_ref:
    nombres_ref = sorted([r["nombre"] for r in LISTA_REFERENCISTAS])
    referencista_sel = st.selectbox("Bibliotecario que firma:", nombres_ref)

st.markdown("---")

# Botón Único de Procesamiento Automático
if st.button("🚀 Generar Certificado desde DSpace", type="primary"):
    if not url_dspace or not cedula:
        st.error("Por favor ingresa la URL de DSpace y el número de cédula.")
    else:
        with st.spinner("Consultando DSpace y construyendo certificado..."):
            datos, error = extraer_metadatos_dspace(url_dspace)
            
            if error:
                st.error(error)
            else:
                ref_info = next(item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel)
                
                payload = {
                    "autor": datos["autor"],
                    "cedula": cedula.strip(),
                    "facultad": datos["facultad"],
                    "carrera": datos["carrera"],
                    "handle": datos["handle"],
                    "nivel": nivel_academico,
                    "ref_nombre": ref_info["nombre"],
                    "ref_cargo": ref_info["cargo"]
                }

                # Guardar el resultado en la sesión activa
                st.session_state["cert_buffer"] = crear_documento_word(payload)
                st.session_state["cert_filename"] = f"Certificado_No_Adeudar_{cedula.strip()}.docx"
                st.session_state["cert_info"] = datos

# Si el certificado ya está generado, mostrar información y botón permanente de descarga
if st.session_state["cert_buffer"] is not None:
    st.success(f"✅ ¡Certificado generado automáticamente para **{st.session_state['cert_info']['autor']}**!")
    st.info(f"**Facultad:** {st.session_state['cert_info']['facultad']} | **Carrera:** {st.session_state['cert_info']['carrera']}")
    
    st.download_button(
        label="📥 DESCARGAR DOCUMENTO WORD (.DOCX)",
        data=st.session_state["cert_buffer"],
        file_name=st.session_state["cert_filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="secondary"
    )
