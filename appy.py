import io
import re
from datetime import datetime
from bs4 import BeautifulSoup
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import requests
import urllib3
import streamlit as st

# Desactivar advertencias de certificados SSL no verificados
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuración de la página web
st.set_page_config(
    page_title="Generador de Certificados - UCuenca", 
    page_icon="📜",
    layout="wide"
)

st.title("📜 Generador de Certificados de No Adeudar")
st.subheader("Centro de Documentación Regional 'Juan Bautista Vázquez'")

# ------------------------------------------------------------------
# LISTA OFICIAL DE REFERENCISTAS
# ------------------------------------------------------------------
LISTA_REFERENCISTAS = [
    {"nombre": "DORIS PATRICIA TENESACA CARDENAS", "cargo": "Bibliotecario 2"},
    {"nombre": "ERIKA ELIZABETH IDROVO SALAZAR", "cargo": "Bibliotecario 2"},
    {"nombre": "ERIKA SOFIA PEÑAFIEL VAZQUEZ", "cargo": "Bibliotecario 2"},
    {"nombre": "FRANCISCO TEODORO ASTUDILLO SAQUINAULA", "cargo": "Bibliotecario 2"},
    {"nombre": "JENNY EULALIA PEREZ MEJIA", "cargo": "Bibliotecario 2"},
    {"nombre": "JHOANNA NOEMI MOGOLLON GUZMAN", "cargo": "Bibliotecario 2"},
    {"nombre": "PAOLA DEL ROCIO AMAYA ARCE", "cargo": "Bibliotecario 2"},
    {"nombre": "PATRICIA MARIBEL DUCHI PESANTEZ", "cargo": "Bibliotecario 2"},
    {"nombre": "WILMAN GONZALO TANDAZO GUEVARA", "cargo": "Bibliotecario 2"},
]

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

if "metadatos_dspace" not in st.session_state:
    st.session_state["metadatos_dspace"] = None

# ------------------------------------------------------------------
# EXTRACCIÓN AVANZADA DE METADATOS
# ------------------------------------------------------------------
def extraer_metadatos_dspace(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8',
        'Referer': 'https://dspace.ucuenca.edu.ec/'
    }
    
    match_handle = re.search(r'(\d+/\d+)', url)
    match_uuid = re.search(r'/items/([a-f0-9\-]+)', url)

    autores = []
    facultad, carrera, handle_final = "", "", url
    errores_log = []

    # 1. BÚSQUEDA VÍA DSPACE REST API
    try:
        if match_uuid:
            api_url = f"https://dspace.ucuenca.edu.ec/server/api/core/items/{match_uuid.group(1)}"
        elif match_handle:
            api_url = f"https://dspace.ucuenca.edu.ec/server/api/discover/search/objects?query=handle:{match_handle.group(1)}"
        else:
            api_url = None

        if api_url:
            resp = requests.get(api_url, headers=headers, timeout=15, verify=False)
            if resp.status_code == 200:
                data = resp.json()
                if "_embedded" in data and "searchResult" in data["_embedded"]:
                    objects = data["_embedded"]["searchResult"]["_embedded"]["objects"]
                    item_data = objects[0]["_embedded"]["indexableObject"] if objects else {}
                else:
                    item_data = data

                metadata = item_data.get("metadata", {})

                campos_autores = ["dc.contributor.author", "dc.creator", "dc.author", "dc.contributor"]
                for campo in campos_autores:
                    entries = metadata.get(campo, [])
                    for entry in entries:
                        val = entry.get("value", "").strip().upper()
                        if val and val not in autores:
                            autores.append(val)

                campos_facultad = ["thesis.degree.grantor", "dc.publisher", "dc.contributor.department"]
                for campo in campos_facultad:
                    if campo in metadata and metadata[campo]:
                        facultad = metadata[campo][0].get("value", "")
                        if facultad: break

                campos_carrera = ["thesis.degree.discipline", "dc.subject", "dc.degree.name"]
                for campo in campos_carrera:
                    if campo in metadata and metadata[campo]:
                        carrera = metadata[campo][0].get("value", "")
                        if carrera: break

                handle_final = metadata.get("dc.identifier.uri", [{}])[0].get("value", url)
            else:
                errores_log.append(f"Respuesta API DSpace HTTP {resp.status_code}")
    except Exception as e:
        errores_log.append(f"Fallo de conexión API: {str(e)}")

    # 2. RESPALDO VÍA HTML SCRAPING
    if not autores:
        try:
            resp_html = requests.get(url, headers=headers, timeout=15, verify=False)
            if resp_html.status_code == 200:
                soup = BeautifulSoup(resp_html.text, 'html.parser')
                meta_authors = soup.find_all('meta', {'name': re.compile(r'DC\.creator|DC\.contributor|citation_author', re.I)})
                for ma in meta_authors:
                    val = ma.get('content', '').strip().upper()
                    if val and val not in autores:
                        autores.append(val)

                if not facultad:
                    meta_pub = soup.find('meta', {'name': re.compile(r'DC\.publisher|citation_publisher', re.I)})
                    if meta_pub and meta_pub.get('content'):
                        facultad = meta_pub['content'].strip()

                meta_uri = soup.find('meta', {'name': re.compile(r'DC\.identifier|citation_abstract_html_url', re.I)})
                if meta_uri and meta_uri.get('content'):
                    handle_final = meta_uri['content'].strip()
            else:
                errores_log.append(f"Respuesta HTML DSpace HTTP {resp_html.status_code}")
        except Exception as e:
            errores_log.append(f"Fallo de conexión HTML: {str(e)}")

    if not autores:
        msg_error = "No se pudieron consultar los metadatos automáticamente. Se habilitó el ingreso manual abajo."
        return None, msg_error

    facultad_clean = facultad.replace("Universidad de Cuenca.", "").replace("Universidad de Cuenca", "").strip()
    if not facultad_clean:
        facultad_clean = "Facultad de Ciencias Químicas"

    carrera_clean = carrera.strip()
    if not carrera_clean:
        carrera_clean = "Bioquímica y Farmacia"

    return {
        "autores": autores,
        "facultad": facultad_clean,
        "carrera": carrera_clean,
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
    
    # Anchos equilibrados: 2.1 pulgadas para el logo y 4.4 para el texto oficial
    cell_left.width = Inches(2.1)
    cell_right.width = Inches(4.4)
    
    cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Logo UCUENCA
    p_logo = cell_left.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(0)
    
    run_logo = p_logo.add_run("UCUENCA")
    run_logo.font.name = 'Arial'
    run_logo.font.size = Pt(22)  # Tamaño adecuado para entrar perfectamente en la celda
    run_logo.font.bold = True
    run_logo.font.color.rgb = RGBColor(15, 43, 91)

    # Bloque de texto
    p_hdr = cell_right.paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hdr.paragraph_format.line_spacing = 1.0
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(0)

    r1 = p_hdr.add_run("FORMATO DE NO ADEUDAR MATERIAL BIBLIOGRÁFICO A LA\n")
    r1.font.bold = True
    r1.font.size = Pt(8.5)
    r1.font.name = 'Arial'

    r2 = p_hdr.add_run("BIBLIOTECA\n")
    r2.font.bold = True
    r2.font.size = Pt(8.5)
    r2.font.name = 'Arial'

    r3 = p_hdr.add_run("UC-CDRJVB-FOR-020\n")
    r3.font.bold = True
    r3.font.size = Pt(8.5)
    r3.font.name = 'Arial'

    r4 = p_hdr.add_run("Página 1 de 1")
    r4.font.bold = False
    r4.font.size = Pt(8.5)
    r4.font.name = 'Arial'

    doc.add_paragraph()

    # 2. TÍTULO
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(24)
    p_titulo.paragraph_format.space_after = Pt(24)
    run_titulo = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
    run_titulo.font.name = 'Arial'
    run_titulo.font.size = Pt(13)
    run_titulo.font.bold = True

    # 3. CUERPO
    p_cuerpo = doc.add_paragraph()
    p_cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cuerpo.paragraph_format.line_spacing = 1.15
    p_cuerpo.paragraph_format.space_after = Pt(24)

    prefix_carrera = "de la Carrera de" if datos["nivel"] == "Pregrado" else "del Programa de"

    p_cuerpo.add_run('El Centro de Documentación Regional "Juan Bautista Vázquez" certifica que ').font.name = 'Arial'
    p_cuerpo.runs[0].font.size = Pt(11)

    r_nombre = p_cuerpo.add_run(f'{datos["autor"]}')
    r_nombre.font.name = 'Arial'
    r_nombre.font.size = Pt(11)
    r_nombre.font.bold = True

    r = p_cuerpo.add_run(', portador de la cédula de ciudadanía No. ')
    r.font.name = 'Arial'
    r.font.size = Pt(11)

    r_cedula = p_cuerpo.add_run(f'{datos["cedula"]}')
    r_cedula.font.name = 'Arial'
    r_cedula.font.size = Pt(11)
    r_cedula.font.bold = True

    r_text = f', estudiante de la {datos["facultad"]} {prefix_carrera} {datos["carrera"]}, no adeuda ningún bien, ni material bibliográfico en esta dependencia.'
    r = p_cuerpo.add_run(r_text)
    r.font.name = 'Arial'
    r.font.size = Pt(11)

    # 4. FECHA
    p_fecha = doc.add_paragraph()
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_fecha.paragraph_format.space_after = Pt(28)
    hoy = datetime.now()
    fecha_texto = f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"
    r_fecha = p_fecha.add_run(fecha_texto)
    r_fecha.font.name = 'Arial'
    r_fecha.font.size = Pt(11)

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
    r_nom.font.name = 'Arial'
    r_nom.font.size = Pt(11)
    r_nom.font.bold = True

    r_cargo = p_firma.add_run(f'{datos["ref_cargo"]}\n')
    r_cargo.font.name = 'Arial'
    r_cargo.font.size = Pt(10)

    r_cdr = p_firma.add_run('CDR "Juan Bautista Vázquez"')
    r_cdr.font.name = 'Arial'
    r_cdr.font.size = Pt(10)

    for _ in range(3):
        doc.add_paragraph()

    # 6. PIE DE PÁGINA
    p_link = doc.add_paragraph()
    p_link.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_link.paragraph_format.space_before = Pt(0)
    p_link.paragraph_format.space_after = Pt(14)

    r_lbl = p_link.add_run("Link: ")
    r_lbl.font.name = 'Arial'
    r_lbl.font.size = Pt(10)
    r_lbl.font.bold = True

    r_handle = p_link.add_run(datos["handle"])
    r_handle.font.name = 'Arial'
    r_handle.font.size = Pt(10)
    r_handle.font.underline = True
    r_handle.font.color.rgb = RGBColor(0, 51, 153)

    p_ver = doc.add_paragraph()
    p_ver.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ver.paragraph_format.space_before = Pt(0)
    p_ver.paragraph_format.space_after = Pt(0)

    r_ver = p_ver.add_run("Version: 2.0")
    r_ver.font.name = 'Arial'
    r_ver.font.size = Pt(9.5)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------------
# INTERFAZ DE USUARIO STREAMLIT
# ------------------------------------------------------------------
st.markdown("### 1. Búsqueda de Trabajo en DSpace")

col_url, col_btn = st.columns([3, 1])
with col_url:
    url_input = st.text_input("URL del Ítem en DSpace:", placeholder="https://dspace.ucuenca.edu.ec/items/85088aa6...")

with col_btn:
    st.write(" ")
    if st.button("🔍 Buscar Metadatos", type="primary"):
        if url_input:
            with st.spinner("Consultando DSpace..."):
                data, err = extraer_metadatos_dspace(url_input)
                if err:
                    st.error(err)
                    st.session_state["metadatos_dspace"] = {
                        "autores": ["APELLIDOS NOMBRES ESTUDIANTE"],
                        "facultad": "Facultad de Ciencias Químicas",
                        "carrera": "Bioquímica y Farmacia",
                        "handle": url_input
                    }
                    st.info("ℹ️ Se habilitaron los campos manuales abajo para ingresar datos directamente.")
                else:
                    st.session_state["metadatos_dspace"] = data
                    st.success(f"¡Metadatos procesados! Se encontraron {len(data['autores'])} autor(es).")

if st.session_state["metadatos_dspace"]:
    data = st.session_state["metadatos_dspace"]
    
    st.markdown("---")
    st.markdown("### 2. Configuración General")
    
    col_fac, col_car, col_niv, col_ref = st.columns(4)
    with col_fac:
        facultad_edit = st.text_input("Facultad:", value=data["facultad"])
    with col_car:
        carrera_edit = st.text_input("Carrera / Programa:", value=data["carrera"])
    with col_niv:
        nivel_sel = st.selectbox("Nivel Académico:", ["Pregrado", "Maestría / Posgrado"])
    with col_ref:
        nombres_ref = sorted([r["nombre"] for r in LISTA_REFERENCISTAS])
        referencista_sel = st.selectbox("Bibliotecario que firma:", nombres_ref)

    st.markdown("---")
    st.markdown(f"### 3. Autores ({len(data['autores'])})")

    ref_info = next(item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel)

    for idx, autor_nombre in enumerate(data["autores"], start=1):
        with st.expander(f"👤 Autor {idx}: {autor_nombre}", expanded=True):
            col_a, col_b, col_c = st.columns([2, 2, 2])
            
            with col_a:
                nombre_autor_final = st.text_input(f"Nombre Estudiante {idx}:", value=autor_nombre, key=f"nom_{idx}")
            with col_b:
                cedula_autor = st.text_input(f"Cédula de Ciudadanía {idx}:", placeholder="Ej: 1401064256", key=f"ced_{idx}")
            with col_c:
                st.write(" ")
                if st.button(f"📄 Generar Certificado #{idx}", key=f"btn_{idx}"):
                    if not cedula_autor:
                        st.warning(f"Por favor ingresa la cédula para {nombre_autor_final}")
                    else:
                        payload = {
                            "autor": nombre_autor_final.strip().upper(),
                            "cedula": cedula_autor.strip(),
                            "facultad": facultad_edit.strip(),
                            "carrera": carrera_edit.strip(),
                            "handle": data["handle"],
                            "nivel": nivel_sel,
                            "ref_nombre": ref_info["nombre"],
                            "ref_cargo": ref_info["cargo"]
                        }
                        
                        buf = crear_documento_word(payload)
                        st.download_button(
                            label=f"📥 Descargar Word (.docx)",
                            data=buf,
                            file_name=f"Certificado_No_Adeudar_{cedula_autor.strip()}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{idx}"
                        )
