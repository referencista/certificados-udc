import io
import re
from datetime import datetime
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import requests
import streamlit as st

# ------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ------------------------------------------------------------------
st.set_page_config(
    page_title="Generador de Certificados - UCuenca",
    page_icon="🎓",
    layout="centered",
)

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

  # 1. Autores: conservando comas, orden exacto (APELLIDOS, NOMBRES) y en MAYÚSCULAS
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

  # 2. Facultad y Carrera (revisando múltiples etiquetas posibles)
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

  # 3. Handle Oficial
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
# GENERADOR DEL DOCUMENTO WORD (.DOCX)
# ------------------------------------------------------------------
def crear_documento_word(datos):
  doc = docx.Document()

  for section in doc.sections:
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

  # Encabezado
  table_header = doc.add_table(rows=1, cols=2)
  table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
  table_header.autofit = False

  cell_left, cell_right = (
      table_header.rows[0].cells[0],
      table_header.rows[0].cells[1],
  )
  cell_left.width, cell_right.width = Inches(2.1), Inches(4.4)
  cell_left.vertical_alignment = cell_right.vertical_alignment = (
      WD_ALIGN_VERTICAL.CENTER
  )

  p_logo = cell_left.paragraphs[0]
  p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
  run_logo = p_logo.add_run("UCUENCA")
  run_logo.font.name, run_logo.font.size, run_logo.font.bold = (
      "Arial",
      Pt(22),
      True,
  )
  run_logo.font.color.rgb = RGBColor(15, 43, 91)

  p_hdr = cell_right.paragraphs[0]
  p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_hdr.paragraph_format.line_spacing = 1.0

  p_hdr.add_run(
      "FORMATO DE NO ADEUDAR MATERIAL BIBLIOGRÁFICO A LA\nBIBLIOTECA\nUC-CDRJVB-FOR-020\nPágina"
      " 1 de 1"
  ).font.size = Pt(8.5)

  doc.add_paragraph()

  # Título
  p_titulo = doc.add_paragraph()
  p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_titulo.paragraph_format.space_before = (
      p_titulo.paragraph_format.space_after
  ) = Pt(24)
  r_tit = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
  r_tit.font.name, r_tit.font.size, r_tit.font.bold = "Arial", Pt(13), True

  # Cuerpo
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

  p_cuerpo.add_run(
      'El Centro de Documentación Regional "Juan Bautista Vázquez" certifica'
      " que "
  ).font.name = "Arial"

  r_nom = p_cuerpo.add_run(f'{datos["autor"]}')
  r_nom.font.bold = True

  p_cuerpo.add_run(", portador(a) de la cédula de ciudadanía No. ")

  # Espacio en blanco listo para escribir a mano en Word
  r_ced = p_cuerpo.add_run("____________________")
  r_ced.font.bold = True

  p_cuerpo.add_run(
      f', estudiante de la {datos["facultad"]} {prefix_carrera}'
      f' {datos["carrera"]}, no adeuda ningún bien, ni material bibliográfico'
      " en esta dependencia."
  )

  # Fecha
  p_fecha = doc.add_paragraph()
  p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  hoy = datetime.now()
  p_fecha.add_run(f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}")

  # Firma
  doc.add_paragraph().paragraph_format.space_after = Pt(30)
  p_atentamente = doc.add_paragraph()
  p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_atentamente.add_run(
      "Atentamente,\n\n________________________________________"
  )

  p_firma = doc.add_paragraph()
  p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
  r_f1 = p_firma.add_run(f'\n{datos["ref_nombre"]}\n')
  r_f1.font.bold = True
  p_firma.add_run(f'{datos["ref_cargo"]}\nCDR "Juan Bautista Vázquez"')

  for _ in range(2):
    doc.add_paragraph()

  # Link Handle
  p_link = doc.add_paragraph()
  p_link.add_run("Link: ").font.bold = True
  r_h = p_link.add_run(datos["handle"])
  r_h.font.underline = True
  r_h.font.color.rgb = RGBColor(0, 51, 153)

  p_ver = doc.add_paragraph()
  p_ver.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_ver.add_run("Version: 2.0").font.size = Pt(9.5)

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# ------------------------------------------------------------------
# INTERFAZ GRÁFICA DE STREAMLIT
# ------------------------------------------------------------------
st.title("🎓 Generador de Certificados de No Adeudar")
st.caption("Universidad de Cuenca - CDR 'Juan Bautista Vázquez'")

url_input = st.text_input(
    "🔗 Ingrese la URL del trabajo en DSpace (Handle o UUID):",
    placeholder="https://dspace.ucuenca.edu.ec/items/...",
)

tipo_estudio = st.radio(
    "Nivel de estudio:",
    ["Pregrado", "Maestría", "Doctorado"],
    horizontal=True,
)

referencista_nombre = st.text_input(
    "✍️ Nombre del Referencista:", value="Doris Elizabeth Aguilar Aguilar"
)
referencista_cargo = st.text_input(
    "💼 Cargo del Referencista:", value="Referencista"
)

if url_input.strip():
  with st.spinner("Extrayendo información desde DSpace..."):
    metadatos = extraer_metadatos_dspace(url_input)

  st.subheader("📝 Datos del Certificado")

  # Muestra y permite editar autores si hay varios
  autores_str = st.text_area(
      "Estudiante(s) (Un autor por línea):",
      value="\n".join(metadatos["autores"]),
  )
  facultad = st.text_input("Facultad:", value=metadatos["facultad"])
  carrera = st.text_input("Carrera / Programa:", value=metadatos["carrera"])
  handle_oficial = st.text_input("Link Oficial:", value=metadatos["handle"])

  autores_lista = [a.strip() for a in autores_str.split("\n") if a.strip()]

  if st.button("📄 Generar Certificado Word"):
    if not autores_lista:
      st.error("Por favor ingresa al menos un estudiante.")
    else:
      for autor in autores_lista:
        payload = {
            "autor": autor,
            "facultad": facultad,
            "carrera": carrera,
            "tipo_estudio": tipo_estudio,
            "ref_nombre": referencista_nombre,
            "ref_cargo": referencista_cargo,
            "handle": handle_oficial,
        }

        buf = crear_documento_word(payload)

        st.success(f"Certificado listo para: {autor}")
        st.download_button(
            label=f"⬇️ Descargar Documento Word ({autor[:20]}...)",
            data=buf,
            file_name=f"Certificado_{autor.replace(',', '').replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
