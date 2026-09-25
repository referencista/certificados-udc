import io
import re
import unicodedata
import zipfile
from datetime import datetime

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import pandas as pd  # <-- lo necesitaremos para leer el archivo Excel en la modalidad Complexivo
import requests
import streamlit as st
import urllib3

# Desactivar advertencias SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTILOS UCUENCA
# ------------------------------------------------------------------
st.set_page_config(
    page_title="UCuenca - Generador de Certificados de No Adeudar",
    page_icon="🎓",
    layout="wide",
)

# Estilos CSS Personalizados estilo Universidad de Cuenca (Centrado)
CSS_UCUENCA = """
<style>
    /* Estilo General */
    body {
        font-family: 'Arial', sans-serif;
        background-color: #f4f6f9;
    }
    
    /* Header Institucional Centrado */
    .uc-header {
        text-align: center;
        background: linear-gradient(90deg, #0F2B5B 0%, #163B7A 100%);
        color: white;
        padding: 20px 30px;
        border-radius: 8px;
        margin-bottom: 25px;
        border-bottom: 4px solid #9E1B32;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    }
    .uc-header h1 {
        color: #ffffff !important;
        font-size: 26px !important;
        font-weight: 700 !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    .uc-header p {
        color: #d1dbe8 !important;
        font-size: 14px !important;
        margin-top: 5px !important;
        margin-bottom: 0 !important;
    }

    /* Botones Principales */
    .stButton>button {
        background-color: #0F2B5B !important;
        color: white !important;
        border-radius: 6px !important;
        font-weight: bold !important;
        border: none !important;
        padding: 10px 20px !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        background-color: #9E1B32 !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2) !important;
        color: white !important;
    }

    /* Botones de Descarga */
    .stDownloadButton>button {
        background-color: #1b6ec2 !important;
        color: white !important;
        border-radius: 6px !important;
        font-weight: bold !important;
        border: none !important;
        transition: all 0.3s ease !important;
    }
    .stDownloadButton>button:hover {
        background-color: #0F2B5B !important;
        color: white !important;
    }

    /* Ajuste de Secciones / Expanders */
    .streamlit-expanderHeader {
        background-color: #ffffff !important;
        border-left: 5px solid #0F2B5B !important;
        border-radius: 4px !important;
        font-weight: bold !important;
        color: #0F2B5B !important;
    }
    
    /* Inputs */
    div[data-baseweb="input"] {
        border-radius: 6px !important;
    }
</style>
"""

st.markdown(CSS_UCUENCA, unsafe_allow_html=True)

# Banner de Encabezado Institucional Centrado con Nombre Oficial
st.markdown(
    """
    <div class="uc-header">
        <h1>UNIVERSIDAD DE CUENCA</h1>
        <p>Centro de Documentación Regional “Juan Bautista Vázquez” (CDR-JBV) &bull; Certificado de No Adeudar</p>
    </div>
    """,
    unsafe_allow_html=True,
)

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

FACULTADES_MAP = [
    (
        "Facultad de Ciencias Agropecuarias",
        [
            "agropecuaria",
            "agropecuarias",
            "agronomía",
            "agronomia",
            "agronómica",
            "agronomica",
            "veterinaria",
            "medicina veterinaria",
            "zootecnia",
        ],
    ),
    (
        "Facultad de Ciencias de la Hospitalidad",
        [
            "hospitalidad",
            "turismo",
            "gastronomía",
            "gastronomia",
            "hotelería",
            "hoteleria",
        ],
    ),
    (
        "Facultad de Arquitectura y Urbanismo",
        [
            "arquitectura",
            "diseño gráfico",
            "diseño de interiores",
            "diseño interior",
            "urbanismo",
        ],
    ),
    (
        "Facultad de Ciencias Económicas y Administrativas",
        [
            "económica",
            "economía",
            "economica",
            "economia",
            "administración",
            "administracion",
            "contabilidad",
            "auditoría",
            "auditoria",
            "mercadotecnia",
            "finanzas",
            "comercio exterior",
            "empresa",
            "empresas",
        ],
    ),
    (
        "Facultad de Ingeniería",
        [
            "ingeniería civil",
            "ingenieria civil",
            "sistemas",
            "computación",
            "computacion",
            "eléctrica",
            "electrica",
            "electrónica",
            "electronica",
            "telecomunicaciones",
            "industrial",
            "carreteras",
            "software",
        ],
    ),
    (
        "Facultad de Ciencias Médicas",
        [
            "médica",
            "medicina",
            "enfermería",
            "enfermeria",
            "fisioterapia",
            "laboratorio clínico",
            "laboratorio clinico",
            "nutrición",
            "nutricion",
            "salud",
            "fonoaudiología",
            "imagenología",
        ],
    ),
    (
        "Facultad de Ciencias Químicas",
        [
            "química",
            "quimica",
            "bioquímica",
            "bioquimica",
            "farmacia",
            "ingeniería química",
            "ingenieria quimica",
            "ingeniería ambiental",
            "ingenieria ambiental",
            "alimentos",
        ],
    ),
    (
        "Facultad de Filosofía, Letras y Ciencias de la Educación",
        [
            "filosofía",
            "filosofia",
            "educación",
            "educacion",
            "comunicación",
            "comunicacion",
            "idiomas",
            "lengua",
            "historia",
            "pedagogía",
            "pedagogia",
            "literatura",
            "inicial",
            "básica",
            "basica",
        ],
    ),
    (
        "Facultad de Jurisprudencia, Ciencias Políticas y Sociales",
        [
            "jurisprudencia",
            "derecho",
            "trabajo social",
            "orientación familiar",
            "orientacion familiar",
            "política",
            "politica",
            "género",
        ],
    ),
    (
        "Facultad de Artes",
        [
            "artes",
            "música",
            "musica",
            "danza",
            "teatro",
            "artes visuales",
            "diseño teatral",
            "escénicas",
        ],
    ),
    (
        "Facultad de Psicología",
        [
            "psicología",
            "psicologia",
            "psicología clínica",
            "psicologia clinica",
            "psicoeducativa",
            "psicoterapia",
        ],
    ),
]


def normalizar_texto(texto):
  if not texto:
    return ""
  texto = unicodedata.normalize("NFD", texto)
  texto = re.sub(r"[\u0300-\u036f]", "", texto)
  return texto.lower()


# ------------------------------------------------------------------
# EXTRACCIÓN Y LIMPIEZA DE METADATOS VÍA API REST DSPACE 7
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
    endpoint = f"/pid/find?id={handle_id}&embed=owningCollection"
    handle_official = f"https://dspace.ucuenca.edu.ec/handle/{handle_id}"
  elif match_uuid:
    uuid_id = match_uuid.group(1)
    endpoint = f"/core/items/{uuid_id}?embed=owningCollection"

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
        "facultad": "Facultad de ",
        "carrera": "",
        "handle": handle_official,
    }

  metadata = data.get("metadata", {})

  def get_all_meta_values(keys_list):
    if isinstance(keys_list, str):
      keys_list = [keys_list]
    res = []
    for key in keys_list:
      items = metadata.get(key, [])
      for item in items:
        v = item.get("value", "").strip()
        if v and v not in res:
          res.append(v)
    return res

  raw_authors = get_all_meta_values(["dc.contributor.author", "dc.creator"])
  autores_formateados = []
  for a in raw_authors:
    a_clean = re.sub(r"\s+", " ", a).strip().upper()
    if a_clean and a_clean not in autores_formateados:
      autores_formateados.append(a_clean)

  titulos = " ".join(get_all_meta_values(["dc.title"]))
  materias = " ".join(
      get_all_meta_values(["dc.subject", "dc.description.abstract"])
  )
  texto_inferencia = normalizar_texto(f"{titulos} {materias}")

  candidatos_facultad = get_all_meta_values([
      "thesis.degree.grantor",
      "dc.publisher",
      "dc.department",
      "dc.contributor.department",
      "dc.publisher.department",
      "dc.degree.grantor",
  ])

  owning_coll_name = ""
  try:
    owning_coll_name = (
        data.get("_embedded", {})
        .get("owningCollection", {})
        .get("name", "")
        .strip()
    )
  except Exception:
    pass

  if owning_coll_name:
    candidatos_facultad.append(owning_coll_name)

  facultad_detectada = ""
  for cand in candidatos_facultad:
    cand_norm = normalizar_texto(cand)
    for fac_oficial, kw_list in FACULTADES_MAP:
      fac_norm = normalizar_texto(fac_oficial)
      if fac_norm in cand_norm or any(kw in cand_norm for kw in kw_list):
        facultad_detectada = fac_oficial
        break
    if facultad_detectada:
      break

  if not facultad_detectada:
    for fac_oficial, kw_list in FACULTADES_MAP:
      if any(kw in texto_inferencia for kw in kw_list):
        facultad_detectada = fac_oficial
        break

  candidatos_carrera = get_all_meta_values([
      "thesis.degree.discipline",
      "thesis.degree.name",
      "dc.degree.discipline",
      "dc.degree.program",
      "dc.subject",
  ])

  if owning_coll_name and owning_coll_name not in candidatos_carrera:
    candidatos_carrera.append(owning_coll_name)

  carrera_detectada = ""
  for cand in candidatos_carrera:
    cand_clean = re.sub(
        r"^(Universidad de Cuenca[\.\,\-]?\s*|Facultad de [^.]+\.\s*)",
        "",
        cand,
        flags=re.I,
    ).strip()
    cand_clean = re.sub(
        r"^(carrera de|programa de|maestría en|doctorado en|msc\.|ing\.|lic\.)\s*",
        "",
        cand_clean,
        flags=re.I,
    ).strip()

    if (
        cand_clean
        and "universidad" not in cand_clean.lower()
        and "facultad" not in cand_clean.lower()
    ):
      carrera_detectada = cand_clean
      break

  if not carrera_detectada and candidatos_carrera:
    carrera_detectada = candidatos_carrera[0]

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
      "facultad": (
          facultad_detectada
          if facultad_detectada
          else "Facultad de Ciencias Químicas"
      ),
      "carrera": carrera_detectada,
      "handle": handle_official,
  }


# ------------------------------------------------------------------
# GENERADOR DEL DOCUMENTO WORD (.DOCX) - FORMATO INTACTO
# ------------------------------------------------------------------
def crear_documento_word(datos):
  doc = docx.Document()

  # Estilo global Arial
  style_normal = doc.styles["Normal"]
  font_normal = style_normal.font
  font_normal.name = "Arial"
  font_normal.size = Pt(11)

  # Márgenes
  for section in doc.sections:
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

  # ENCABEZADO
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

  p_logo = cell_left.paragraphs[0]
  p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
  p_logo.paragraph_format.space_after = Pt(0)
  run_logo = p_logo.add_run("UCUENCA")
  run_logo.font.name = "Arial"
  run_logo.font.size = Pt(28)
  run_logo.font.bold = True
  run_logo.font.color.rgb = RGBColor(15, 43, 91)

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
  doc.add_paragraph()

  # TÍTULO PRINCIPAL
  p_titulo = doc.add_paragraph()
  p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_titulo.paragraph_format.space_before = Pt(24)
  p_titulo.paragraph_format.space_after = Pt(24)
  
    

  r_tit = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
  r_tit.font.name = "Arial"
  r_tit.font.size = Pt(13)
  r_tit.font.bold = True

  # CUERPO DEL CERTIFICADO (Se mantiene exacto a la redacción original)
  p_cuerpo = doc.add_paragraph()
  p_cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
  p_cuerpo.paragraph_format.line_spacing = 1.5
  p_cuerpo.paragraph_format.space_after = Pt(24)

  fac_clean = re.sub(
      r"^(Universidad de Cuenca[\.\,\-]?\s*)", "", datos["facultad"], flags=re.I
  ).strip()
  if not fac_clean.lower().startswith("facultad de") and fac_clean:
    fac_clean = f"Facultad de {fac_clean}"

  carr_clean = re.sub(
      r"^(Universidad de Cuenca[\.\,\-]?\s*)", "", datos["carrera"], flags=re.I
  ).strip()
  carr_clean = re.sub(
      r"^(carrera de|programa de|maestría en|doctorado en)\s*",
      "",
      carr_clean,
      flags=re.I,
  ).strip()

  tipo = datos.get("tipo_estudio", "Pregrado")

  r_c1 = p_cuerpo.add_run(
      'El Centro de Documentación Regional "Juan Bautista Vázquez" certifica'
      " que "
  )
  r_c1.font.name = "Arial"

  r_nom = p_cuerpo.add_run(f'{datos["autor"]}')
  r_nom.font.name = "Arial"
  r_nom.font.bold = True

  r_c2 = p_cuerpo.add_run(", portador de la cédula de ciudadanía No. ")
  r_c2.font.name = "Arial"

  r_ced = p_cuerpo.add_run("____________________")
  r_ced.font.name = "Arial"
  r_ced.font.bold = True

  # Redacción según modalidad
  if tipo == "Complexivo":
      texto_cert = (
          f", estudiante de la {fac_clean}, de la {carr_clean}, "
          "modalidad Examen Complexivo, no adeuda ningún bien, ni material bibliográfico en esta dependencia."
      )
  else:
      if tipo == "Maestría":
          prefix_carrera = "del Programa de Maestría en"
      elif tipo == "Doctorado":
          prefix_carrera = "del Programa de Doctorado en"
      else:
          prefix_carrera = "de la Carrera de"
      
      texto_cert = ( 
          f", estudiante de la {fac_clean} {prefix_carrera} {carr_clean}, "
                "no adeuda ningún bien, ni material bibliográfico en esta dependencia."
      )
  r_c3 = p_cuerpo.add_run(texto_cert)
  r_c3.font.name = "Arial"

  # FECHA
  p_fecha = doc.add_paragraph()
  p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_fecha.paragraph_format.space_after = Pt(24)

  hoy = datetime.now()
  r_fecha = p_fecha.add_run(
      f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"
  )
  r_fecha.font.name = "Arial"

  # FIRMAS
  doc.add_paragraph().paragraph_format.space_after = Pt(30)

  p_atentamente = doc.add_paragraph()
  p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
  r_at = p_atentamente.add_run(
      "Atentamente,\n\n\n\n\n\n________________________________________"
  )
  r_at.font.name = "Arial"

  p_firma = doc.add_paragraph()
  p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER

  r_f1 = p_firma.add_run(f'\n{datos["ref_nombre"]}\n')
  r_f1.font.name = "Arial"
  r_f1.font.bold = True

  r_f2 = p_firma.add_run(f'{datos["ref_cargo"]}\nCDR "Juan Bautista Vázquez"')
  r_f2.font.name = "Arial"

  # ENLACE DSPACE
  for _ in range(2):
    doc.add_paragraph()

  p_link = doc.add_paragraph()
  r_l1 = p_link.add_run("Link / Registro: " if tipo == "Complexivo" else "Link: ")
  r_l1.font.name = "Arial"
  r_l1.font.bold = True

  valor_link = "Examen Complexivo / Registro Interno" if tipo == "Complexivo" else datos.get("handle", "")
  r_h = p_link.add_run(valor_link)
  r_h.font.name = "Arial"
  r_h.font.underline = True
  r_h.font.color.rgb = RGBColor(0, 51, 153)

  # VERSIÓN
  doc.add_paragraph()
  doc.add_paragraph()

  p_ver = doc.add_paragraph()
  p_ver.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  r_v = p_ver.add_run("Version: 1.0")
  r_v.font.name = "Arial"
  r_v.font.size = Pt(9.5)

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer

# ------------------------------------------------------------------
# INTERFAZ PRINCIPAL EN STREAMLIT
# ------------------------------------------------------------------

# 1. Selector principal de Tipo de Titulación y Referencista
col_tipo, col_ref = st.columns([2, 2])

with col_tipo:
    tipo_estudio = st.selectbox(
        "Tipo de Titulación / Proceso:",
        ["Pregrado", "Maestría", "Doctorado", "Complexivo"]
    )

with col_ref:
    nombres_ref = sorted([r["nombre"] for r in LISTA_REFERENCISTAS])
    referencista_sel = st.selectbox("Referencista que firma:", nombres_ref)

ref_info = next(
    item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel
)

st.markdown("---")

# ------------------------------------------------------------------
# CASO A: PREGRADO, MAESTRÍA O DOCTORADO (DSPACE)
# ------------------------------------------------------------------
if tipo_estudio != "Complexivo":
    st.markdown("#### 1. Parámetros de la Consulta DSpace")
    
    url_input = st.text_input(
        "Handle de DSpace:",
        placeholder="Ej: https://dspace.ucuenca.edu.ec/handle/123456789/49197",
    )

    col_btn1, col_btn2 = st.columns([3, 1])

    with col_btn1:
        btn_procesar = st.button(
            "🔍 Consultar DSpace y Generar Documento(s)", use_container_width=True
        )

    with col_btn2:
        btn_limpiar = st.button(
            "🧹 Nueva Búsqueda", use_container_width=True
        )

    if btn_limpiar:
        st.session_state.pop("datos_cargados", None)
        st.rerun()

    if btn_procesar or "datos_cargados" in st.session_state:
        if btn_procesar:
            if not url_input:
                st.warning("⚠️ Por favor ingresa el Handle de DSpace.")
                st.stop()
            st.session_state.pop("datos_cargados", None)
            
            with st.spinner("Conectando con el repositorio DSpace de la UCuenca..."):
                meta = extraer_metadatos_dspace(url_input)
                st.session_state["datos_cargados"] = meta

        meta = st.session_state["datos_cargados"]

        st.markdown("---")
        st.markdown("#### 2. Validación de Metadatos Extramunicipales y Estudiantes")

        col_f, col_c = st.columns(2)
        with col_f:
            facultad_final = st.text_input("Facultad Detectada:", value=meta["facultad"])
        with col_c:
            carrera_final = st.text_input(
                "Carrera / Programa Detectado:", value=meta["carrera"]
            )

        st.markdown(f"**Estudiantes / Autores encontrados ({len(meta['autores'])})**")

        certificados_generados = []

        for idx, autor_nombre in enumerate(meta["autores"], start=1):
            with st.expander(
                f"🎓 Estudiante #{idx}: {autor_nombre}", expanded=True
            ):
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

                buf = crear_documento_word(payload)
                st.download_button(
                    label=f"📄 Descargar Certificado Word - Estudiante {idx}",
                    data=buf,
                    file_name=f"Certificado_{nom_est.replace(' ', '_')}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                    key=f"btn_dl_{idx}",
                )

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
                    "📦 Descargar TODOS los Certificados en Un Archivo ZIP"
                    f" ({len(certificados_generados)} archivos)"
                ),
                data=zip_buffer,
                file_name="Certificados_No_Adeudar_UCuenca.zip",
                mime="application/zip",
            )

# ------------------------------------------------------------------
# CASO B: EXAMEN COMPLEXIVO
# ------------------------------------------------------------------
else:
    st.markdown("#### 1. Parámetros de Examen Complexivo")
    
    modo_complexivo = st.radio(
        "Seleccione la modalidad de generación:",
        ["👤 Ingreso Manual (1 Estudiante)", "📂 Carga Masiva desde Excel (Listado)"],
        horizontal=True
    )

    st.markdown("---")

    if "Ingreso Manual" in modo_complexivo:
        st.info("ℹ️ Ingrese los datos del estudiante de examen complexivo para generar su certificado individual.")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            nom_comp = st.text_input("Nombres y Apellidos Completos:", value="ZHIZHPON GUACHICHULLCA BYRON GEOVANNY")
            carr_comp = st.text_input("Carrera o Programa:", value="PEDAGOGIA DE LAS ARTES Y LAS HUMANIDADES")
        with col_m2:
            fac_comp = st.text_input("Facultad:", value="FILOSOFIA, LETRAS Y CIENCIAS DE LA EDUCACION")

        payload_manual = {
            "autor": nom_comp.strip().upper(),
            "facultad": fac_comp.strip(),
            "carrera": carr_comp.strip(),
            "tipo_estudio": "Complexivo",
            "handle": "",
            "ref_nombre": ref_info["nombre"],
            "ref_cargo": ref_info["cargo"],
        }

        buf_manual = crear_documento_word(payload_manual)
        
        st.download_button(
            label="📄 Descargar Certificado Word (.docx)",
            data=buf_manual,
            file_name=f"Certificado_Complexivo_{nom_comp.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_complexivo_manual"
        )

    else:
        st.info("📂 Carga de Excel en preparación (Paso 3)...")
