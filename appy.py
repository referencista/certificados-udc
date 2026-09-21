import io
import re
import zipfile
from datetime import datetime
from bs4 import BeautifulSoup
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import requests
import urllib3
import streamlit as st

# Desactivar advertencias SSL de DSpace
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

# Mapeo inteligente de Facultades
FACULTADES_MAP = [
    ("Facultad de Arquitectura y Urbanismo", ["arquitectura", "diseño gráfico", "diseño de interiores", "diseño interior"]),
    ("Facultad de Ciencias Agropecuarias", ["agropecuaria", "agronomía", "agronomia", "veterinaria", "medicina veterinaria"]),
    ("Facultad de Ciencias Económicas y Administrativas", ["económica", "economía", "administración", "contabilidad", "auditoría", "mercadotecnia", "finanzas", "comercio exterior"]),
    ("Facultad de Ingeniería", ["ingeniería civil", "sistemas", "computación", "eléctrica", "electrónica", "telecomunicaciones", "industrial"]),
    ("Facultad de Ciencias Médicas", ["médica", "medicina", "enfermería", "fisioterapia", "laboratorio clínico", "nutrición", "fonoaudiología", "estimulación temprana"]),
    ("Facultad de Ciencias Químicas", ["química", "bioquímica", "farmacia", "ingeniería química", "ingeniería ambiental"]),
    ("Facultad de Filosofía, Letras y Ciencias de la Educación", ["filosofía", "educación", "comunicación", "idiomas", "lengua", "historia", "geografía", "matemáticas", "física", "pedagogía"]),
    ("Facultad de Jurisprudencia, Ciencias Políticas y Sociales", ["jurisprudencia", "derecho", "trabajo social", "orientación familiar", "ciencias políticas"]),
    ("Facultad de Artes", ["artes", "música", "danza", "teatro", "artes visuales", "artes escénicas"]),
    ("Facultad de Psicología", ["psicología", "psicología clínica", "psicología educativa"])
]

# ------------------------------------------------------------------
# EXTRACCIÓN AVANZADA DE METADATOS DSPACE UCUENCA
# ------------------------------------------------------------------
def extraer_metadatos_dspace(url_input):
    url_clean = url_input.strip()
    match_handle = re.search(r'(\d+/\d+)', url_clean)
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8'
    }

    urls_to_try = []
    handle_official = ""
    
    if match_handle:
        h = match_handle.group(1)
        handle_official = f"http://dspace.ucuenca.edu.ec/handle/{h}"
        urls_to_try.append(f"https://dspace.ucuenca.edu.ec/handle/{h}?mode=full")
        urls_to_try.append(f"https://dspace.ucuenca.edu.ec/handle/{h}")
    elif url_clean.startswith("http"):
        handle_official = url_clean
        urls_to_try.append(url_clean)
        if "?mode=full" not in url_clean:
            urls_to_try.insert(0, url_clean + ("&mode=full" if "?" in url_clean else "?mode=full"))
    else:
        return None, "Ingrese una URL válida de DSpace o un código Handle (ej. 123456789/40123)."

    html_content = ""
    for u in urls_to_try:
        try:
            resp = requests.get(u, headers=headers, timeout=12, verify=False)
            if resp.status_code == 200 and len(resp.text) > 500:
                html_content = resp.text
                break
        except Exception:
            continue

    if not html_content:
        return None, "No se pudo consultar el repositorio de DSpace. Verifique que la URL o Handle sea correcto."

    soup = BeautifulSoup(html_content, 'html.parser')

    # 1. AUTORES
    raw_authors = []

    # Opción A: Meta tags
    meta_authors = soup.find_all('meta', {'name': re.compile(r'^(DC\.creator|citation_author|dc\.contributor\.author)$', re.I)})
    for ma in meta_authors:
        val = ma.get('content', '').strip()
        if val and val not in raw_authors:
            raw_authors.append(val)

    # Opción B: Tabla en ?mode=full
    for row in soup.find_all('tr'):
        tds = row.find_all(['td', 'th'])
        if len(tds) >= 2:
            lbl = tds[0].get_text().strip().lower()
            val = tds[1].get_text().strip()
            if ('dc.contributor.author' in lbl or 'dc.creator' in lbl) and val:
                if val not in raw_authors:
                    raw_authors.append(val)

    # Reordenar y formatear nombres de autores
    autores_formateados = []
    for a in raw_authors:
        a_clean = re.sub(r'\s+', ' ', a).strip()
        if ',' in a_clean:
            parts = a_clean.split(',', 1)
            apellidos = parts[0].strip()
            nombres = parts[1].strip()
            nombre_completo = f"{nombres} {apellidos}".upper()
        else:
            nombre_completo = a_clean.upper()
        
        if nombre_completo not in autores_formateados:
            autores_formateados.append(nombre_completo)

    # 2. FACULTAD Y CARRERA
    raw_grantor = ""
    raw_discipline = ""
    raw_publisher = ""
    raw_subjects = []

    for row in soup.find_all('tr'):
        tds = row.find_all(['td', 'th'])
        if len(tds) >= 2:
            lbl = tds[0].get_text().strip().lower()
            val = tds[1].get_text().strip()
            if 'thesis.degree.grantor' in lbl or 'dc.publisher' in lbl:
                if not raw_grantor: raw_grantor = val
            elif 'thesis.degree.discipline' in lbl:
                if not raw_discipline: raw_discipline = val
            elif 'dc.subject' in lbl and val:
                raw_subjects.append(val)

    # Colecciones y migas de pan
    collections = soup.find_all(['a', 'span', 'li'], class_=re.compile(r'breadcrumb|collection|trail', re.I))
    coll_text = " ".join([c.get_text().strip() for c in collections if c.get_text().strip()])

    body_text = soup.get_text()

    # Nivel Académico
    full_str = f"{coll_text} {raw_discipline} {' '.join(raw_subjects)} {body_text}".lower()
    nivel = "Maestría / Posgrado" if any(w in full_str for w in ["posgrado", "maestría", "maestria", "magíster", "master"]) else "Pregrado"

    # Determinar Facultad
    facultad_detectada = ""
    match_fac = re.search(r'Facultad de [A-Za-zÁÉÍÓÚáéíóúñÑ\s,]+', f"{raw_grantor} {coll_text} {body_text}")
    if match_fac:
        candidate = match_fac.group(0).split('\n')[0].split('-')[0].strip()
        candidate = re.sub(r'[.\,\;]+$', '', candidate)
        if len(candidate) < 65:
            facultad_detectada = candidate

    if not facultad_detectada:
        search_target = f"{raw_grantor} {raw_publisher} {coll_text} {raw_discipline} {' '.join(raw_subjects)}".lower()
        for fac_nombre, keywords in FACULTADES_MAP:
            if any(kw in search_target for kw in keywords):
                facultad_detectada = fac_nombre
                break

    if not facultad_detectada:
        facultad_detectada = "Facultad de Ciencias Químicas"

    # Determinar Carrera
    carrera_detectada = raw_discipline.strip() if raw_discipline else ""

    if not carrera_detectada:
        match_coll = re.search(r'([A-Za-zÁÉÍÓÚáéíóúñÑ\s]+)\s*-\s*(Pregrado|Posgrado|Tesis|Maestría)', f"{coll_text} {body_text}")
        if match_coll:
            cand = match_coll.group(1).strip()
            if len(cand) > 3 and "Collections" not in cand:
                carrera_detectada = cand

    if not carrera_detectada:
        match_car = re.search(r'Carrera de ([A-Za-zÁÉÍÓÚáéíóúñÑ\s]+)', f"{coll_text} {body_text}")
        if match_car:
            carrera_detectada = match_car.group(1).strip().split('\n')[0]

    if not carrera_detectada and raw_subjects:
        carrera_detectada = raw_subjects[0]

    if not carrera_detectada:
        carrera_detectada = "Bioquímica y Farmacia"

    return {
        "autores": autores_formateados if autores_formateados else ["ESTUDIANTE APELLIDOS NOMBRES"],
        "facultad": facultad_detectada.replace("Universidad de Cuenca.", "").strip(),
        "carrera": carrera_detectada.strip(),
        "nivel": nivel,
        "handle": handle_official or url_clean
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

    # ENCABEZADO
    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False

    cell_left = table_header.rows[0].cells[0]
    cell_right = table_header.rows[0].cells[1]
    
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
    run_logo.font.size = Pt(22)
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

    # TÍTULO
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(24)
    p_titulo.paragraph_format.space_after = Pt(24)
    run_titulo = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
    run_titulo.font.name = 'Arial'
    run_titulo.font.size = Pt(13)
    run_titulo.font.bold = True

    # CUERPO
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

    # FECHA
    p_fecha = doc.add_paragraph()
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_fecha.paragraph_format.space_after = Pt(28)
    hoy = datetime.now()
    fecha_texto = f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"
    r_fecha = p_fecha.add_run(fecha_texto)
    r_fecha.font.name = 'Arial'
    r_fecha.font.size = Pt(11)

    # FIRMA
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

    # PIE DE PÁGINA
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
st.markdown("### 1. Búsqueda por Enlace o Handle de DSpace")

col_url, col_btn = st.columns([3, 1])
with col_url:
    url_input = st.text_input("URL o Handle de DSpace:", placeholder="Ej: https://dspace.ucuenca.edu.ec/handle/123456789/40123")

with col_btn:
    st.write(" ")
    if st.button("🔍 Extraer Metadatos", type="primary"):
        if url_input:
            with st.spinner("Obteniendo información oficial de DSpace..."):
                data, err = extraer_metadatos_dspace(url_input)
                if err:
                    st.error(err)
                else:
                    st.session_state["metadatos_dspace"] = data
                    st.success(f"✅ ¡Éxito! Se detectó la **{data['facultad']}** - **{data['carrera']}** y **{len(data['autores'])} autor(es)**.")

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
        nivel_index = 0 if data["nivel"] == "Pregrado" else 1
        nivel_sel = st.selectbox("Nivel Académico:", ["Pregrado", "Maestría / Posgrado"], index=nivel_index)
    with col_ref:
        nombres_ref = sorted([r["nombre"] for r in LISTA_REFERENCISTAS])
        referencista_sel = st.selectbox("Bibliotecario que firma:", nombres_ref)

    st.markdown("---")
    st.markdown(f"### 3. Estudiantes / Autores Encontrados ({len(data['autores'])})")

    ref_info = next(item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel)

    datos_autores_para_descarga = []

    for idx, autor_nombre in enumerate(data["autores"], start=1):
        with st.expander(f"👤 Autor #{idx}: {autor_nombre}", expanded=True):
            col_a, col_b = st.columns([3, 2])
            
            with col_a:
                nom_final = st.text_input(f"Nombre Estudiante {idx}:", value=autor_nombre, key=f"nom_{idx}")
            with col_b:
                ced_final = st.text_input(f"Cédula de Ciudadanía {idx}:", placeholder="Ej: 0101234567", key=f"ced_{idx}")

            payload_indiv = {
                "autor": nom_final.strip().upper(),
                "cedula": ced_final.strip(),
                "facultad": facultad_edit.strip(),
                "carrera": carrera_edit.strip(),
                "handle": data["handle"],
                "nivel": nivel_sel,
                "ref_nombre": ref_info["nombre"],
                "ref_cargo": ref_info["cargo"]
            }
            datos_autores_para_descarga.append(payload_indiv)

            if ced_final.strip():
                buf = crear_documento_word(payload_indiv)
                st.download_button(
                    label=f"📄 Descargar Certificado de {nom_final.split()[0]}",
                    data=buf,
                    file_name=f"Certificado_{ced_final.strip()}_{nom_final.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_indiv_{idx}"
                )
            else:
                st.info("💡 Ingrese el número de cédula para habilitar la descarga individual.")

    # SI HAY MÚLTIPLES AUTORES: Descarga en lote .ZIP
    if len(datos_autores_para_descarga) > 1:
        st.markdown("---")
        st.markdown("### 4. Descarga Conjunta")
        
        # Verificar que todos tengan cédula
        cedulas_completas = all(a["cedula"] for a in datos_autores_para_descarga)
        
        if cedulas_completas:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for a_data in datos_autores_para_descarga:
                    doc_buf = crear_documento_word(a_data)
                    file_name = f"Certificado_{a_data['cedula']}_{a_data['autor'].replace(' ', '_')}.docx"
                    zf.writestr(file_name, doc_buf.getvalue())
            
            zip_buffer.seek(0)
            st.download_button(
                label=f"📦 Descargar TODOS los Certificados ({len(datos_autores_para_descarga)} archivos en .ZIP)",
                data=zip_buffer,
                file_name="Certificados_No_Adeudar_UCuenca.zip",
                mime="application/zip",
                type="primary"
            )
        else:
            st.warning("⚠️ Para generar el paquete .ZIP con todos los certificados, asegúrese de ingresar las cédulas de todos los autores.")
