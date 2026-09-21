import io
import re
import zipfile
import unicodedata
from datetime import datetime
from bs4 import BeautifulSoup
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import requests
import urllib3
import streamlit as st

# Desactivar advertencias SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuración de la interfaz Streamlit
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

# Mapa exhaustivo de facultades y palabras clave
FACULTADES_MAP = [
    ("Facultad de Ciencias Agropecuarias", ["agropecuaria", "agropecuarias", "agronomía", "agronomia", "agronómica", "agronomica", "veterinaria", "medicina veterinaria"]),
    ("Facultad de Ciencias de la Hospitalidad", ["hospitalidad", "turismo", "gastronomía", "gastronomia", "hotelería", "hoteleria"]),
    ("Facultad de Arquitectura y Urbanismo", ["arquitectura", "diseño gráfico", "diseño de interiores", "diseño interior"]),
    ("Facultad de Ciencias Económicas y Administrativas", ["económica", "economía", "administración", "contabilidad", "auditoría", "mercadotecnia", "finanzas", "comercio exterior"]),
    ("Facultad de Ingeniería", ["ingeniería civil", "sistemas", "computación", "eléctrica", "electrónica", "telecomunicaciones", "industrial"]),
    ("Facultad de Ciencias Médicas", ["médica", "medicina", "enfermería", "fisioterapia", "laboratorio clínico", "nutrición"]),
    ("Facultad de Ciencias Químicas", ["química", "bioquímica", "farmacia", "ingeniería química", "ingeniería ambiental"]),
    ("Facultad de Filosofía, Letras y Ciencias de la Educación", ["filosofía", "educación", "comunicación", "idiomas", "lengua", "historia", "pedagogía"]),
    ("Facultad de Jurisprudencia, Ciencias Políticas y Sociales", ["jurisprudencia", "derecho", "trabajo social", "orientación familiar"]),
    ("Facultad de Artes", ["artes", "música", "danza", "teatro"]),
    ("Facultad de Psicología", ["psicología", "psicología clínica"])
]

def normalizar_texto(texto):
    if not texto: return ""
    texto = unicodedata.normalize('NFD', texto)
    texto = re.sub(r'[\u0300-\u036f]', '', texto)
    return texto.lower()

# ------------------------------------------------------------------
# EXTRACCIÓN AVANZADA DE METADATOS DSPACE CON CACHÉ DE MEMORIA RAM
# ------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def extraer_metadatos_dspace(url_input):
    url_clean = url_input.strip()
    match_handle = re.search(r'(\d+/\d+)', url_clean)
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8'
    }

    if url_clean.startswith("http"):
        handle_official = url_clean.split("?")[0]
    elif match_handle:
        handle_official = f"https://dspace.ucuenca.edu.ec/handle/{match_handle.group(1)}"
    else:
        handle_official = url_clean

    urls_to_try = []
    if match_handle:
        h = match_handle.group(1)
        urls_to_try.append(f"https://dspace.ucuenca.edu.ec/handle/{h}?mode=full")
        urls_to_try.append(f"https://dspace.ucuenca.edu.ec/handle/{h}")
    elif url_clean.startswith("http"):
        urls_to_try.append(url_clean)
        if "?mode=full" not in url_clean:
            urls_to_try.insert(0, url_clean + ("&mode=full" if "?" in url_clean else "?mode=full"))

    html_content = ""
    for u in urls_to_try:
        try:
            resp = requests.get(u, headers=headers, timeout=10, verify=False)
            if resp.status_code == 200 and len(resp.text) > 400:
                html_content = resp.text
                break
        except Exception:
            continue

    raw_authors, facultad_detectada, carrera_detectada = [], "", ""

    if html_content:
        soup = BeautifulSoup(html_content, 'html.parser')

        # 1. Autores
        meta_authors = soup.find_all('meta', {'name': re.compile(r'^(DC\.creator|citation_author|dc\.contributor\.author)$', re.I)})
        for ma in meta_authors:
            val = ma.get('content', '').strip()
            if val and val not in raw_authors:
                raw_authors.append(val)

        for row in soup.find_all('tr'):
            tds = row.find_all(['td', 'th'])
            if len(tds) >= 2:
                lbl = tds[0].get_text().strip().lower()
                val = tds[1].get_text().strip()
                if ('dc.contributor.author' in lbl or 'dc.creator' in lbl) and val:
                    if val not in raw_authors:
                        raw_authors.append(val)

        # 2. Extraer campos estructurados
        raw_grantor, raw_discipline = "", ""
        for row in soup.find_all('tr'):
            tds = row.find_all(['td', 'th'])
            if len(tds) >= 2:
                lbl = tds[0].get_text().strip().lower()
                val = tds[1].get_text().strip()
                if 'thesis.degree.grantor' in lbl or 'dc.publisher' in lbl:
                    if not raw_grantor: raw_grantor = val
                elif 'thesis.degree.discipline' in lbl:
                    if not raw_discipline: raw_discipline = val

        full_text_norm = normalizar_texto(soup.get_text())

        # Detección de Facultad
        match_fac = re.search(r'facultad de [a-záéíóúñ\s,]+', f"{raw_grantor.lower()} {soup.get_text().lower()}")
        if match_fac:
            cand = match_fac.group(0).split('\n')[0].split('-')[0].strip().title()
            if len(cand) < 65 and "Facultad" in cand:
                facultad_detectada = re.sub(r'[.\,\;]+$', '', cand)

        if not facultad_detectada:
            for fac_nombre, keywords in FACULTADES_MAP:
                if any(kw in full_text_norm for kw in keywords):
                    facultad_detectada = fac_nombre
                    break

        # Detección de Carrera
        if raw_discipline:
            carrera_detectada = raw_discipline.strip()
        else:
            if "agronomia" in full_text_norm or "agronomica" in full_text_norm:
                carrera_detectada = "Ingeniería Agronómica"
            elif "veterinaria" in full_text_norm:
                carrera_detectada = "Medicina Veterinaria"
            elif "turismo" in full_text_norm:
                carrera_detectada = "Turismo"
            elif "gastronomia" in full_text_norm:
                carrera_detectada = "Gastronomía"

    # Formatear nombres de autores (APELLIDOS NOMBRES -> NOMBRES APELLIDOS)
    autores_formateados = []
    for a in raw_authors:
        a_clean = re.sub(r'\s+', ' ', a).strip()
        if ',' in a_clean:
            parts = a_clean.split(',', 1)
            nombre_completo = f"{parts[1].strip()} {parts[0].strip()}".upper()
        else:
            nombre_completo = a_clean.upper()
        if nombre_completo not in autores_formateados:
            autores_formateados.append(nombre_completo)

    return {
        "autores": autores_formateados if autores_formateados else ["APELLIDOS NOMBRES ESTUDIANTE"],
        "facultad": facultad_detectada if facultad_detectada else "Facultad de Ciencias Agropecuarias",
        "carrera": carrera_detectada if carrera_detectada else "Ingeniería Agronómica",
        "handle": handle_official
    }

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

    # ENCABEZADO CON LOGO
    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False

    cell_left = table_header.rows[0].cells[0]
    cell_right = table_header.rows[0].cells[1]
    
    cell_left.width = Inches(2.1)
    cell_right.width = Inches(4.4)
    
    cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_logo = cell_left.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(0)
    
    run_logo = p_logo.add_run("UCUENCA")
    run_logo.font.name = 'Arial'
    run_logo.font.size = Pt(22)
    run_logo.font.bold = True
    run_logo.font.color.rgb = RGBColor(15, 43, 91)

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

    # CUERPO DEL TEXTO
    p_cuerpo = doc.add_paragraph()
    p_cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cuerpo.paragraph_format.line_spacing = 1.15
    p_cuerpo.paragraph_format.space_after = Pt(24)

    tipo = datos.get("tipo_estudio", "Pregrado")
    if tipo == "Maestría":
        prefix_carrera = "del Programa de Maestría en"
    elif tipo == "Doctorado":
        prefix_carrera = "del Programa de Doctorado en"
    else:  # Pregrado o Complexivo
        prefix_carrera = "de la Carrera de"

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
# INTERFAZ PRINCIPAL STREAMLIT
# ------------------------------------------------------------------
st.markdown("### 1. Parámetros de la Consulta")

col1, col2, col3 = st.columns([3, 2, 2])

with col1:
    url_input = st.text_input("URL o Handle de DSpace:", placeholder="Ej: https://dspace.ucuenca.edu.ec/handle/123456789/49197")

with col2:
    tipo_estudio = st.selectbox("Tipo de Titulación:", ["Pregrado", "Maestría", "Doctorado", "Complexivo"])

with col3:
    nombres_ref = sorted([r["nombre"] for r in LISTA_REFERENCISTAS])
    referencista_sel = st.selectbox("Referencista que firma:", nombres_ref)

btn_procesar = st.button("🔍 Extraer y Preparar Certificado(s)", type="primary", use_container_width=True)

if btn_procesar or "datos_cargados" in st.session_state:
    if btn_procesar:
        if not url_input:
            st.warning("⚠️ Por favor ingresa la URL o Handle de DSpace.")
            st.stop()
        
        with st.spinner("Procesando información de DSpace..."):
            meta = extraer_metadatos_dspace(url_input)
            st.session_state["datos_cargados"] = meta

    meta = st.session_state["datos_cargados"]
    ref_info = next(item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel)

    st.markdown("---")
    st.markdown("### 2. Confirmación de Metadatos y Estudiantes")

    col_f, col_c = st.columns(2)
    with col_f:
        facultad_final = st.text_input("Facultad:", value=meta["facultad"])
    with col_c:
        carrera_final = st.text_input("Carrera / Programa:", value=meta["carrera"])

    st.markdown(f"#### Autores Detectados ({len(meta['autores'])})")

    certificados_generados = []

    for idx, autor_nombre in enumerate(meta["autores"], start=1):
        with st.expander(f"👤 Estudiante #{idx}: {autor_nombre}", expanded=True):
            col_a, col_b = st.columns([3, 2])
            
            with col_a:
                nom_est = st.text_input(f"Nombre Estudiante #{idx}:", value=autor_nombre, key=f"nom_{idx}")
            with col_b:
                ced_est = st.text_input(f"Cédula de Ciudadanía #{idx}:", placeholder="Ej: 0101234567", key=f"ced_{idx}")

            payload = {
                "autor": nom_est.strip().upper(),
                "cedula": ced_est.strip(),
                "facultad": facultad_final.strip(),
                "carrera": carrera_final.strip(),
                "tipo_estudio": tipo_estudio,
                "handle": meta["handle"],
                "ref_nombre": ref_info["nombre"],
                "ref_cargo": ref_info["cargo"]
            }
            certificados_generados.append(payload)

            if ced_est.strip():
                buf = crear_documento_word(payload)
                st.download_button(
                    label=f"📥 Descargar Certificado Word (.docx) - Estudiante {idx}",
                    data=buf,
                    file_name=f"Certificado_{ced_est.strip()}_{nom_est.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"btn_dl_{idx}"
                )
            else:
                st.info("💡 Ingrese el número de cédula para activar el botón de descarga.")

    if len(certificados_generados) > 1:
        st.markdown("---")
        if all(c["cedula"] for c in certificados_generados):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for c_data in certificados_generados:
                    doc_buf = crear_documento_word(c_data)
                    zf.writestr(f"Certificado_{c_data['cedula']}_{c_data['autor'].replace(' ', '_')}.docx", doc_buf.getvalue())
            
            zip_buffer.seek(0)
            st.download_button(
                label=f"📦 Descargar TODOS los Certificados ({len(certificados_generados)} archivos .ZIP)",
                data=zip_buffer,
                file_name="Certificados_No_Adeudar_UCuenca.zip",
                mime="application/zip",
                type="primary"
            )
        else:
            st.warning("⚠️ Complete las cédulas de todos los estudiantes para descargar el paquete .ZIP conjunto.")
