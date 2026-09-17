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

st.title("📜 Generador de Certificado de No Adeudar")
st.subheader("Centro de Documentación Regional 'Juan Bautista Vázquez'")

# ------------------------------------------------------------------
# DICCIONARIOS Y LISTAS DE CONFIGURACIÓN
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

# ------------------------------------------------------------------
# FUNCIÓN PARA EXTRAER METADATOS DE DSPACE 7+
# ------------------------------------------------------------------
def obtener_metadatos_dspace(url):
    try:
        match = re.search(r'/items/([a-f0-9\-]+)', url)
        if not match:
            match = re.search(r'/handle/(\d+/\d+)', url)
            if not match:
                return None, "No se encontró un UUID o Handle válido en la URL."
            handle_str = match.group(1)
            api_url = f"https://dspace.ucuenca.edu.ec/server/api/discover/search/objects?query=handle:{handle_str}"
        else:
            uuid = match.group(1)
            api_url = f"https://dspace.ucuenca.edu.ec/server/api/core/items/{uuid}"

        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(api_url, headers=headers, timeout=10)
        
        if response.status_code != 200:
            return None, f"Error de conexión con DSpace (Código: {response.status_code})."

        data = response.json()
        if "_embedded" in data and "searchResult" in data["_embedded"]:
            objects = data["_embedded"]["searchResult"]["_embedded"]["objects"]
            if objects:
                item_data = objects[0]["_embedded"]["indexableObject"]
            else:
                return None, "No se encontraron metadatos para esta URL."
        else:
            item_data = data

        metadata = item_data.get("metadata", {})

        autor = metadata.get("dc.contributor.author", [{}])[0].get("value", "")
        facultad = metadata.get("thesis.degree.grantor", [{}])[0].get("value", "")
        carrera = metadata.get("thesis.degree.discipline", [{}])[0].get("value", "")
        handle = metadata.get("dc.identifier.uri", [{}])[0].get("value", url)

        # Formatear el nombre en MAYÚSCULAS
        autor = autor.upper().strip()

        return {
            "autor": autor,
            "facultad": facultad,
            "carrera": carrera,
            "handle": handle
        }, None

    except Exception as e:
        return None, f"Excepción al consultar DSpace: {str(e)}"

# ------------------------------------------------------------------
# FUNCIÓN PARA GENERAR EL DOCUMENTO WORD EXACTO (.DOCX)
# ------------------------------------------------------------------
def crear_documento_word(datos):
    doc = docx.Document()

    # Configuración de márgenes (2.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # 1. ENCABEZADO CON TABLA (Logo UCUENCA a la izquierda / Formato a la derecha)
    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False

    cell_left = table_header.rows[0].cells[0]
    cell_right = table_header.rows[0].cells[1]
    cell_left.width = Inches(3.2)
    cell_right.width = Inches(3.3)

    # Lado Izquierdo: Logotipo UCUENCA
    p_logo = cell_left.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_logo = p_logo.add_run("UCUENCA")
    run_logo.font.name = 'Arial'
    run_logo.font.size = Pt(28)
    run_logo.font.bold = True
    run_logo.font.color.rgb = RGBColor(15, 43, 91) # Azul institucional

    # Lado Derecho: Metadatos del Formulario
    p_hdr = cell_right.paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    r1 = p_hdr.add_run("FORMATO DE NO ADEUDAR MATERIAL BIBLIOGRÁFICO A LA BIBLIOTECA\n")
    r1.font.bold = True
    r1.font.size = Pt(8.5)
    r1.font.name = 'Arial'

    r2 = p_hdr.add_run("UC-CDRJVB-FOR-020\n")
    r2.font.bold = True
    r2.font.size = Pt(8.5)
    r2.font.name = 'Arial'

    r3 = p_hdr.add_run("Página 1 de 1")
    r3.font.size = Pt(8.5)
    r3.font.name = 'Arial'

    doc.add_paragraph() # Espaciador

    # 2. TÍTULO CENTRADO
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_before = Pt(24)
    p_titulo.paragraph_format.space_after = Pt(24)

    run_titulo = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
    run_titulo.font.name = 'Arial'
    run_titulo.font.size = Pt(13)
    run_titulo.font.bold = True

    # 3. CUERPO DEL CERTIFICADO (JUSTIFICADO)
    p_cuerpo = doc.add_paragraph()
    p_cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cuerpo.paragraph_format.line_spacing = 1.15
    p_cuerpo.paragraph_format.space_after = Pt(24)

    prefix_carrera = "de la Carrera de" if datos["nivel"] == "Pregrado" else "del Programa de"

    # Construcción incremental con negritas según plantilla oficial
    r = p_cuerpo.add_run('El Centro de Documentación Regional "Juan Bautista Vázquez" certifica que ')
    r.font.name = 'Arial'; r.font.size = Pt(11)

    r_nombre = p_cuerpo.add_run(f'{datos["autor"]}')
    r_nombre.font.name = 'Arial'; r_nombre.font.size = Pt(11); r_nombre.font.bold = True

    r = p_cuerpo.add_run(', portador de la cédula de ciudadanía No. ')
    r.font.name = 'Arial'; r.font.size = Pt(11)

    r_cedula = p_cuerpo.add_run(f'{datos["cedula"]}')
    r_cedula.font.name = 'Arial'; r_cedula.font.size = Pt(11); r_cedula.font.bold = True

    r_text = f', estudiante de la {datos["facultad"]} {prefix_carrera} {datos["carrera"]}, no adeuda ningún bien, ni material bibliográfico en esta dependencia.'
    r = p_cuerpo.add_run(r_text)
    r.font.name = 'Arial'; r.font.size = Pt(11)

    # 4. FECHA (ALINEADA A LA DERECHA)
    p_fecha = doc.add_paragraph()
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_fecha.paragraph_format.space_after = Pt(28)

    hoy = datetime.now()
    fecha_texto = f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"
    r_fecha = p_fecha.add_run(fecha_texto)
    r_fecha.font.name = 'Arial'; r_fecha.font.size = Pt(11)

    # 5. ATENTAMENTE Y FIRMA
    p_atentamente = doc.add_paragraph()
    p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_atentamente.paragraph_format.space_after = Pt(40)
    r_at = p_atentamente.add_run("Atentamente,")
    r_at.font.name = 'Arial'; r_at.font.size = Pt(11)

    # Línea de Firma
    p_linea = doc.add_paragraph()
    p_linea.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_linea.paragraph_format.space_after = Pt(4)
    r_linea = p_linea.add_run("________________________________________")
    r_linea.font.name = 'Arial'

    # Datos del Bibliotecario
    p_firma = doc.add_paragraph()
    p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_firma.paragraph_format.line_spacing = 1.1

    r_nom = p_firma.add_run(f'{datos["ref_nombre"]}\n')
    r_nom.font.name = 'Arial'; r_nom.font.size = Pt(11); r_nom.font.bold = True

    r_cargo = p_firma.add_run(f'{datos["ref_cargo"]}\n')
    r_cargo.font.name = 'Arial'; r_cargo.font.size = Pt(10)

    r_cdr = p_firma.add_run('CDR "Juan Bautista Vázquez"')
    r_cdr.font.name = 'Arial'; r_cdr.font.size = Pt(10)

    # Espacio previo al pie
    for _ in range(3):
        doc.add_paragraph()

    # 6. PIE DE PÁGINA (LINK A LA IZQUIERDA / VERSIÓN A LA DERECHA)
    table_footer = doc.add_table(rows=1, cols=2)
    table_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_footer.autofit = False

    cell_f_left = table_footer.rows[0].cells[0]
    cell_f_right = table_footer.rows[0].cells[1]
    cell_f_left.width = Inches(5.0)
    cell_f_right.width = Inches(1.5)

    p_link = cell_f_left.paragraphs[0]
    p_link.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_l1 = p_link.add_run("Link: ")
    r_l1.font.name = 'Arial'; r_l1.font.size = Pt(9.5)
    
    r_l2 = p_link.add_run(datos["handle"])
    r_l2.font.name = 'Arial'; r_l2.font.size = Pt(9.5); r_l2.font.underline = True
    r_l2.font.color.rgb = RGBColor(0, 51, 153)

    p_ver = cell_f_right.paragraphs[0]
    p_ver.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_v = p_ver.add_run("Version: 2.0")
    r_v.font.name = 'Arial'; r_v.font.size = Pt(8.5)

    # Guardar en buffer de memoria
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------------
# INTERFAZ WEB EN STREAMLIT
# ------------------------------------------------------------------
st.markdown("---")

col1, col2 = st.columns([2, 1])

with col1:
    url_dspace = st.text_input(
        "URL del Ítem en DSpace:", 
        placeholder="Ejemplo: https://dspace.ucuenca.edu.ec/items/8c34b565-..."
    )

with col2:
    cedula = st.text_input("Número de Cédula:", placeholder="1401064256")

col_lvl, col_ref = st.columns(2)

with col_lvl:
    nivel_academico = st.selectbox("Nivel Académico:", ["Pregrado", "Maestría / Posgrado"])

with col_ref:
    nombres_referencistas = sorted([r["nombre"] for r in LISTA_REFERENCISTAS])
    referencista_sel = st.selectbox("Bibliotecario que firma:", nombres_referencistas)

st.markdown("---")

if st.button("🚀 Generar Certificado (.docx)", type="primary"):
    if not url_dspace or not cedula:
        st.error("Por favor completa la URL de DSpace y la cédula de ciudadanía.")
    else:
        with st.spinner("Consultando DSpace y construyendo el certificado..."):
            datos_dspace, error = obtener_metadatos_dspace(url_dspace)
            
            if error and not datos_dspace:
                st.warning(f"⚠️ {error}")
                st.info("Ingresa los datos manualmente para continuar sin DSpace:")
                
                with st.form("form_manual"):
                    autor_m = st.text_input("Nombre Completo (APELLIDOS NOMBRES):")
                    facultad_m = st.text_input("Facultad:", value="Facultad de Ciencias Químicas")
                    carrera_m = st.text_input("Carrera / Programa:", value="Bioquímica y Farmacia")
                    submit_manual = st.form_submit_button("Generar con Datos Manuales")
                    
                    if submit_manual:
                        datos_dspace = {
                            "autor": autor_m.upper(),
                            "facultad": facultad_m,
                            "carrera": carrera_m,
                            "handle": url_dspace
                        }

            if datos_dspace:
                ref_info = next(item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel)
                
                payload = {
                    "autor": datos_dspace["autor"],
                    "cedula": cedula,
                    "facultad": datos_dspace["facultad"],
                    "carrera": datos_dspace["carrera"],
                    "handle": datos_dspace["handle"],
                    "nivel": nivel_academico,
                    "ref_nombre": ref_info["nombre"],
                    "ref_cargo": ref_info["cargo"]
                }

                docx_buffer = crear_documento_word(payload)

                st.success("¡Certificado generado exitosamente!")
                st.download_button(
                    label="📥 Descargar Documento Word (.docx)",
                    data=docx_buffer,
                    file_name=f"Certificado_No_Adeudar_{cedula}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
