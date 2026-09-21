import io
import re
import unicodedata
import zipfile
from datetime import datetime
from bs4 import BeautifulSoup
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import requests
import streamlit as st
import urllib3

# Desactivar advertencias SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuración de la interfaz Streamlit
st.set_page_config(
    page_title="Generador de Certificados - UCuenca",
    page_icon="📜",
    layout="wide",
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
    {
        "nombre": "FRANCISCO TEODORO ASTUDILLO SAQUINAULA",
        "cargo": "Bibliotecario 2",
    },
    {"nombre": "JENNY EULALIA PEREZ MEJIA", "cargo": "Bibliotecario 2"},
    {"nombre": "JHOANNA NOEMI MOGOLLON GUZMAN", "cargo": "Bibliotecario 2"},
    {"nombre": "PAOLA DEL ROCIO AMAYA ARCE", "cargo": "Bibliotecario 2"},
    {"nombre": "PATRICIA MARIBEL DUCHI PESANTEZ", "cargo": "Bibliotecario 2"},
    {"nombre": "WILMAN GONZALO TANDAZO GUEVARA", "cargo": "Bibliotecario 2"},
]

MESES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]


def normalizar_texto(texto):
  if not texto:
    return ""
  texto = unicodedata.normalize("NFD", texto)
  texto = re.sub(r"[\u0300-\u036f]", "", texto)
  return texto.lower()


# ------------------------------------------------------------------
# EXTRACCIÓN AVANZADA DE METADATOS VÍA API REST DSPACE 7 (UCUENCA)
# ------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def extraer_metadatos_dspace(url_input):
  url_clean = url_input.strip()

  match_handle = re.search(r"(\d+/\d+)", url_clean)
  match_uuid = re.search(
      r"([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})",
      url_clean,
      re.I,
  )

  base_apis = [
      "https://rest-dspace.ucuenca.edu.ec/server/api",
      "https://dspace.ucuenca.edu.ec/server/api",
  ]

  endpoint = None
  handle_official = url_clean

  if match_handle:
    handle_id = match_handle.group(1)
    endpoint = f"/pid/find?id={handle_id}"
    handle_official = f"https://dspace.ucuenca.edu.ec/handle/{handle_id}"
  elif match_uuid:
    uuid_id = match_uuid.group(1)
    endpoint = f"/core/items/{uuid_id}"

  headers = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}
  data = None

  if endpoint:
    for base in base_apis:
      try:
        resp = requests.get(
            base + endpoint, headers=headers, timeout=10, verify=False
        )
        if resp.status_code == 200:
          data = resp.json()
          break
      except Exception:
        continue

  if not data:
    return {
        "autores": ["APELLIDOS, NOMBRES ESTUDIANTE"],
        "facultad": "",
        "carrera": "",
        "handle": handle_official,
    }

  metadata = data.get("metadata", {})

  def get_meta_value(keys_list, default=""):
    if isinstance(keys_list, str):
      keys_list = [keys_list]
    for key in keys_list:
      items = metadata.get(key, [])
      if items and len(items) > 0:
        val = items[0].get("value", "").strip()
        if val:
          return val
    return default

  # Autores
  raw_authors = [
      a.get("value")
      for a in metadata.get("dc.contributor.author", [])
      if a.get("value")
  ]
  autores_formateados = []
  for a in raw_authors:
    a_clean = re.sub(r"\s+", " ", a).strip().upper()
    if a_clean and a_clean not in autores_formateados:
      autores_formateados.append(a_clean)

  # Facultad y Carrera
  facultad = get_meta_value(
      [
          "thesis.degree.grantor",
          "dc.publisher",
          "dc.department",
          "dc.contributor.department",
      ],
      default="",
  )

  carrera = get_meta_value(
      [
          "thesis.degree.discipline",
          "thesis.degree.name",
          "dc.degree.discipline",
          "dc.subject",
      ],
      default="",
  )

  # Handle oficial
  uri_items = metadata.get("dc.identifier.uri", [])
  for uri in uri_items:
    val_uri = uri.get("value", "")
    if "handle/" in val_uri:
      match_h = re.search(r"(\d+/\d+)", val_uri)
      if match_h:
        handle_official = (
            f"https://dspace.ucuenca.edu.ec/handle/{match_h.group(1)}"
        )
        break

  return {
      "autores": (
          autores_formateados
          if autores_formateados
          else ["APELLIDOS, NOMBRES ESTUDIANTE"]
      ),
      "facultad": facultad,
      "carrera": carrera,
      "handle": handle_official,
  }


# ------------------------------------------------------------------
# GENERADOR DEL DOCUMENTO WORD (.DOCX) - TODO EN ARIAL
# ------------------------------------------------------------------
def crear_documento_word(datos):
  doc = docx.Document()

  # Configuración global del estilo base a Arial
  style_normal = doc.styles["Normal"]
  font_normal = style_normal.font
  font_normal.name = "Arial"
  font_normal.size = Pt(11)

  # Márgenes de la página
  for section in doc.sections:
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

  # ------------------------------------------------------------------
  # ENCABEZADO: Separa el logo de la Universidad del texto normativo
  # ------------------------------------------------------------------
  table_header = doc.add_table(rows=1, cols=2)
  table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
  table_header.autofit = False

  cell_left, cell_right = (
      table_header.rows[0].cells[0],
      table_header.rows[0].cells[1],
  )
  cell_left.width = Inches(2.6)
  cell_right.width = Inches(3.9)
  cell_left.vertical_alignment = cell_right.vertical_alignment = (
      WD_ALIGN_VERTICAL.CENTER
  )

  # Celda Izquierda: Logo Grande
  p_logo = cell_left.paragraphs[0]
  p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
  p_logo.paragraph_format.space_after = Pt(0)
  run_logo = p_logo.add_run("UCUENCA")
  run_logo.font.name = "Arial"
  run_logo.font.size = Pt(28)  # Logo más grande
  run_logo.font.bold = True
  run_logo.font.color.rgb = RGBColor(15, 43, 91)

  # Celda Derecha: Texto Normativo
  p_hdr = cell_right.paragraphs[0]
  p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_hdr.paragraph_format.line_spacing = 1.15
  p_hdr.paragraph_format.space_after = Pt(0)

  run_hdr = p_hdr.add_run(
      "FORMATO DE NO ADEUDAR MATERIAL BIBLIOGRÁFICO A LA BIBLIOTECA\n"
      "UC-CDRJVB-FOR-020\n"
      "Página 1 de 1"
  )
  run_hdr.font.name = "Arial"
  run_hdr.font.size = Pt(8.5)

  doc.add_paragraph()

  # ------------------------------------------------------------------
  # TÍTULO PRINCIPAL
  # ------------------------------------------------------------------
  p_titulo = doc.add_paragraph()
  p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_titulo.paragraph_format.space_before = Pt(24)
  p_titulo.paragraph_format.space_after = Pt(24)

  r_tit = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
  r_tit.font.name = "Arial"
  r_tit.font.size = Pt(13)
  r_tit.font.bold = True

  # ------------------------------------------------------------------
  # CUERPO DEL CERTIFICADO
  # ------------------------------------------------------------------
  p_cuerpo = doc.add_paragraph()
  p_cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
  p_cuerpo.paragraph_format.line_spacing = 1.15
  p_cuerpo.paragraph_format.space_after = Pt(24)

  tipo = datos.get("tipo_estudio", "Pregrado")
  prefix_carrera = (
      "del Programa de Maestría en"
      if tipo == "Maestría"
      else (
          "del Programa de Doctorado en"
          if tipo == "Doctorado"
          else "de la Carrera de"
      )
  )

  r_c1 = p_cuerpo.add_run(
      'El Centro de Documentación Regional "Juan Bautista Vázquez" certifica'
      " que "
  )
  r_c1.font.name = "Arial"

  r_nom = p_cuerpo.add_run(f'{datos["autor"]}')
  r_nom.font.name = "Arial"
  r_nom.font.bold = True

  r_c2 = p_cuerpo.add_run(", portador(a) de la cédula de ciudadanía No. ")
  r_c2.font.name = "Arial"

  r_ced = p_cuerpo.add_run("____________________")
  r_ced.font.name = "Arial"
  r_ced.font.bold = True

  r_c3 = p_cuerpo.add_run(
      f', estudiante de la {datos["facultad"]} {prefix_carrera}'
      f' {datos["carrera"]}, no adeuda ningún bien, ni material bibliográfico'
      " en esta dependencia."
  )
  r_c3.font.name = "Arial"

  # ------------------------------------------------------------------
  # FECHA
  # ------------------------------------------------------------------
  p_fecha = doc.add_paragraph()
  p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_fecha.paragraph_format.space_after = Pt(24)

  hoy = datetime.now()
  r_fecha = p_fecha.add_run(
      f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"
  )
  r_fecha.font.name = "Arial"

  # ------------------------------------------------------------------
  # FIRMAS
  # ------------------------------------------------------------------
  doc.add_paragraph().paragraph_format.space_after = Pt(30)

  p_atentamente = doc.add_paragraph()
  p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
  r_at = p_atentamente.add_run(
      "Atentamente,\n\n________________________________________"
  )
  r_at.font.name = "Arial"

  p_firma = doc.add_paragraph()
  p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER

  r_f1 = p_firma.add_run(f'\n{datos["ref_nombre"]}\n')
  r_f1.font.name = "Arial"
  r_f1.font.bold = True

  r_f2 = p_firma.add_run(f'{datos["ref_cargo"]}\nCDR "Juan Bautista Vázquez"')
  r_f2.font.name = "Arial"

  # ------------------------------------------------------------------
  # ENLACE DSPACE
  # ------------------------------------------------------------------
  for _ in range(2):
    doc.add_paragraph()

  p_link = doc.add_paragraph()
  r_l1 = p_link.add_run("Link: ")
  r_l1.font.name = "Arial"
  r_l1.font.bold = True

  r_h = p_link.add_run(datos["handle"])
  r_h.font.name = "Arial"
  r_h.font.underline = True
  r_h.font.color.rgb = RGBColor(0, 51, 153)

  # ------------------------------------------------------------------
  # VERSIÓN (Desplazada dos espacios más abajo)
  # ------------------------------------------------------------------
  doc.add_paragraph()
  doc.add_paragraph()

  p_ver = doc.add_paragraph()
  p_ver.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  r_v = p_ver.add_run("Version: 2.0")
  r_v.font.name = "Arial"
  r_v.font.size = Pt(9.5)

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
  url_input = st.text_input(
      "URL o Handle de DSpace:",
      placeholder="Ej: https://dspace.ucuenca.edu.ec/handle/123456789/49197",
  )

with col2:
  tipo_estudio = st.selectbox(
      "Tipo de Titulación:", ["Pregrado", "Maestría", "Doctorado", "Complexivo"]
  )

with col3:
  nombres_ref = sorted([r["nombre"] for r in LISTA_REFERENCISTAS])
  referencista_sel = st.selectbox("Referencista que firma:", nombres_ref)

btn_procesar = st.button(
    "🔍 Extraer y Preparar Certificado(s)",
    type="primary",
    use_container_width=True,
)

if btn_procesar or "datos_cargados" in st.session_state:
  if btn_procesar:
    if not url_input:
      st.warning("⚠️ Por favor ingresa la URL o Handle de DSpace.")
      st.stop()

    with st.spinner("Procesando información de DSpace..."):
      meta = extraer_metadatos_dspace(url_input)
      st.session_state["datos_cargados"] = meta

  meta = st.session_state["datos_cargados"]
  ref_info = next(
      item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel
  )

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
      nom_est = st.text_input(
          f"Nombre Estudiante #{idx}:", value=autor_nombre, key=f"nom_{idx}"
      )

      payload = {
          "autor": nom_est.strip().upper(),
          "facultad": facultad_final.strip(),
          "carrera": carrera_final.strip(),
          "tipo_estudio": tipo_estudio,
          "handle": meta["handle"],
          "ref_nombre": ref_info["nombre"],
          "ref_cargo": ref_info["cargo"],
      }
      certificados_generados.append(payload)

      # Descarga directa individual
      buf = crear_documento_word(payload)
      st.download_button(
          label=(
              f"📥 Descargar Certificado Word (.docx) - Estudiante {idx}"
          ),
          data=buf,
          file_name=f"Certificado_{nom_est.replace(' ', '_')}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
          key=f"btn_dl_{idx}",
      )

  # Descarga conjunto en ZIP si hay múltiples estudiantes
  if len(certificados_generados) > 1:
    st.markdown("---")
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
      for c_data in certificados_generados:
        doc_buf = crear_documento_word(c_data)
        zf.writestr(
            f"Certificado_{c_data['autor'].replace(' ', '_')}.docx",
            doc_buf.getvalue(),
        )

    zip_buffer.seek(0)
    st.download_button(
        label=(
            "📦 Descargar TODOS los Certificados"
            f" ({len(certificados_generados)} archivos .ZIP)"
        ),
        data=zip_buffer,
        file_name="Certificados_No_Adeudar_UCuenca.zip",
        mime="application/zip",
        type="primary",
    )
