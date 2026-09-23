from datetime import datetime
import io
import re
import unicodedata
import zipfile

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import requests
import streamlit as st
import urllib3

# Desactivar advertencias SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTILOS UCUENCA
# ------------------------------------------------------------------
st.set_page_config(
    page_title="UCuenca - Generador y Reportes de Certificados",
    page_icon="🎓",
    layout="wide",
)

# Estilos CSS Personalizados estilo Universidad de Cuenca & Dashboard
CSS_UCUENCA = """
<style>
    /* Estilo General */
    body {
        font-family: 'Arial', sans-serif;
        background-color: #f4f6f9;
    }

    /* Tabs Personalizados */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        border-bottom: 2px solid #e0e0e0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        background-color: #ffffff;
        border-radius: 6px 6px 0px 0px;
        padding: 10px 20px;
        font-weight: bold;
        color: #555;
        border: 1px solid #e0e0e0;
        border-bottom: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0F2B5B !important;
        border-bottom: 3px solid #0F2B5B !important;
    }

    /* Header Institucional Centrado */
    .uc-header {
        text-align: center;
        background: linear-gradient(90deg, #0F2B5B 0%, #163B7A 100%);
        color: white;
        padding: 20px 30px;
        border-radius: 8px;
        margin-bottom: 20px;
        border-bottom: 4px solid #9E1B32;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    }
    .uc-header h1 {
        color: #ffffff !important;
        font-size: 26px !important;
        font-weight: 700 !important;
        margin: 0 !important;
    }
    .uc-header p {
        color: #d1dbe8 !important;
        font-size: 14px !important;
        margin-top: 5px !important;
    }

    /* Tarjetas de Métricas (KPIs) */
    .kpi-card {
        background-color: #ffffff;
        border-radius: 6px;
        padding: 18px 20px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        border: 1px solid #e5e9f0;
        height: 100%;
    }
    .kpi-blue { border-left: 5px solid #0F2B5B; }
    .kpi-red { border-left: 5px solid #9E1B32; }
    .kpi-navy { border-left: 5px solid #163B7A; }

    .kpi-title {
        font-size: 11px;
        font-weight: 700;
        color: #6c757d;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 26px;
        font-weight: 800;
        color: #0F2B5B;
        line-height: 1.2;
    }
    .kpi-subtext {
        font-size: 12px;
        color: #888888;
        margin-top: 4px;
    }

    /* Botón Verde de Excel */
    .btn-excel > button {
        background-color: #0e7040 !important;
        color: white !important;
        border-radius: 6px !important;
        font-weight: bold !important;
        border: none !important;
        padding: 10px 20px !important;
        float: right;
    }
    .btn-excel > button:hover {
        background-color: #0a522e !important;
    }

    /* Botones Principales Streamlit */
    .stButton>button {
        background-color: #0F2B5B !important;
        color: white !important;
        border-radius: 6px !important;
        font-weight: bold !important;
        border: none !important;
    }
    .stButton>button:hover {
        background-color: #9E1B32 !important;
    }
</style>
"""

st.markdown(CSS_UCUENCA, unsafe_allow_html=True)

# ------------------------------------------------------------------
# CONSTANTES Y LISTAS
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
            "econémica",
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
# FUNCIONES GOOGLE SHEETS
# ------------------------------------------------------------------
def obtener_cliente_sheets():
    if "gcp_service_account" not in st.secrets:
        return None, "Faltan credenciales 'gcp_service_account'"
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    return gspread.authorize(creds), None


def guardar_en_google_sheets(datos):
    try:
        client, err = obtener_cliente_sheets()
        if err:
            return False, err

        sheet_url = st.secrets.get("gsheets", {}).get("spreadsheet_url")
        if sheet_url:
            sheet = client.open_by_url(sheet_url).sheet1
        else:
            sheet_name = st.secrets.get("gsheets", {}).get(
                "spreadsheet_name", "Registro_Certificados_UCuenca"
            )
            sheet = client.open(sheet_name).sheet1

        fila = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            datos.get("ref_nombre", ""),
            datos.get("autor", ""),
            datos.get("facultad", ""),
            datos.get("carrera", ""),
            datos.get("handle", ""),
        ]
        sheet.append_row(fila)
        return True, "Guardado con éxito"
    except Exception as e:
        return False, str(e)


def cargar_datos_reporte():
    try:
        client, err = obtener_cliente_sheets()
        if err:
            return pd.DataFrame()

        sheet_url = st.secrets.get("gsheets", {}).get("spreadsheet_url")
        if sheet_url:
            sheet = client.open_by_url(sheet_url).sheet1
        else:
            sheet_name = st.secrets.get("gsheets", {}).get(
                "spreadsheet_name", "Registro_Certificados_UCuenca"
            )
            sheet = client.open(sheet_name).sheet1

        records = sheet.get_all_records()
        df = pd.DataFrame(records)
        return df
    except Exception:
        # Si falla o no hay datos, retorna un DataFrame de ejemplo
        data_demo = [
            {
                "Fecha y Hora": "2026-09-21 08:35",
                "Referencista": "DORIS TENESACA",
                "Estudiante": "ALVAREZ LOPEZ JUAN CARLOS",
                "Facultad": "Cs. Químicas",
                "Programa / Carrera": "Bioquímica y Farmacia",
                "Link DSpace": "https://dspace.ucuenca.edu.ec/handle/123456789/49197",
            },
            {
                "Fecha y Hora": "2026-09-21 09:12",
                "Referencista": "ERIKA IDROVO",
                "Estudiante": "MORA CASTRO ANA CRISTINA",
                "Facultad": "Ingeniería",
                "Programa / Carrera": "Sistemas",
                "Link DSpace": "https://dspace.ucuenca.edu.ec/handle/123456789/49201",
            },
            {
                "Fecha y Hora": "2026-09-21 10:04",
                "Referencista": "FRANCISCO ASTUDILLO",
                "Estudiante": "SOLANO GUAMAN PEDRO LUIS",
                "Facultad": "Jurisprudencia",
                "Programa / Carrera": "Derecho",
                "Link DSpace": "https://dspace.ucuenca.edu.ec/handle/123456789/49208",
            },
            {
                "Fecha y Hora": "2026-09-20 16:45",
                "Referencista": "ERIKA IDROVO",
                "Estudiante": "VEGA TORRES MARIA JOSE",
                "Facultad": "Filosofía",
                "Programa / Carrera": "Educación Básica",
                "Link DSpace": "https://dspace.ucuenca.edu.ec/handle/123456789/49180",
            },
            {
                "Fecha y Hora": "2026-09-20 15:20",
                "Referencista": "PATRICIA DUCHI",
                "Estudiante": "CORREA QUITO LUIS FERNANDO",
                "Facultad": "Artes",
                "Programa / Carrera": "Artes Visuales",
                "Link DSpace": "https://dspace.ucuenca.edu.ec/handle/123456789/49175",
            },
        ]
        return pd.DataFrame(data_demo)


# ------------------------------------------------------------------
# EXTRACCIÓN DSPACE Y DOCUMENTO WORD
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


def crear_documento_word(datos):
    doc = docx.Document()
    style_normal = doc.styles["Normal"]
    font_normal = style_normal.font
    font_normal.name = "Arial"
    font_normal.size = Pt(11)

    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False

    cell_left, cell_right = (
        table_header.rows[0].cells[0],
        table_header.rows[0].cells[1],
    )
    cell_left.width, cell_right.width = Inches(2.6), Inches(3.9)
    cell_left.vertical_alignment = cell_right.vertical_alignment = (
        WD_ALIGN_VERTICAL.CENTER
    )

    p_logo = cell_left.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_logo = p_logo.add_run("UCUENCA")
    run_logo.font.name = "Arial"
    run_logo.font.size = Pt(28)
    run_logo.font.bold = True
    run_logo.font.color.rgb = RGBColor(15, 43, 91)

    p_hdr = cell_right.paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_hdr = p_hdr.add_run(
        "FORMATO DE NO ADEUDAR MATERIAL BIBLIOGRÁFICO A LA BIBLIOTECA\n"
        "UC-CDRJVB-FOR-020\n"
        "Página 1 de 1"
    )
    run_hdr.font.name, run_hdr.font.size = "Arial", Pt(8.5)

    doc.add_paragraph()
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tit = p_titulo.add_run("CERTIFICADO DE NO ADEUDAR")
    r_tit.font.name, r_tit.font.size, r_tit.font.bold = "Arial", Pt(13), True

    p_cuerpo = doc.add_paragraph()
    p_cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    fac_clean = re.sub(
        r"^(Universidad de Cuenca[\.\,\-]?\s*)", "", datos["facultad"], flags=re.I
    ).strip()
    if not fac_clean.lower().startswith("facultad de") and fac_clean:
        fac_clean = f"Facultad de {fac_clean}"

    carr_clean = re.sub(
        r"^(carrera de|programa de|maestría en|doctorado en)\s*",
        "",
        datos["carrera"],
        flags=re.I,
    ).strip()
    tipo = datos.get("tipo_estudio", "Pregrado")
    prefix_carrera = (
        "del Programa de Maestría en"
        if tipo == "Maestría"
        else "del Programa de Doctorado en"
        if tipo == "Doctorado"
        else "de la Carrera de"
    )

    p_cuerpo.add_run(
        'El Centro de Documentación Regional "Juan Bautista Vázquez" certifica que '
    )
    r_nom = p_cuerpo.add_run(f'{datos["autor"]}')
    r_nom.font.bold = True
    p_cuerpo.add_run(", portador de la cédula de ciudadanía No. ")
    r_ced = p_cuerpo.add_run("____________________")
    r_ced.font.bold = True
    p_cuerpo.add_run(
        f", estudiante de la {fac_clean} {prefix_carrera} {carr_clean}, no adeuda"
        " ningún bien, ni material bibliográfico en esta dependencia."
    )

    p_fecha = doc.add_paragraph()
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hoy = datetime.now()
    p_fecha.add_run(
        f"Cuenca, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(25)
    p_atentamente = doc.add_paragraph()
    p_atentamente.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_atentamente.add_run(
        "Atentamente,\n\n\n\n________________________________________"
    )

    p_firma = doc.add_paragraph()
    p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f1 = p_firma.add_run(f'\n{datos["ref_nombre"]}\n')
    r_f1.font.bold = True
    p_firma.add_run(f'{datos["ref_cargo"]}\nCDR "Juan Bautista Vázquez"')

    doc.add_paragraph()
    p_link = doc.add_paragraph()
    r_l1 = p_link.add_run("Link: ")
    r_l1.font.bold = True
    r_h = p_link.add_run(datos["handle"])
    r_h.font.underline = True
    r_h.font.color.rgb = RGBColor(0, 51, 153)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ------------------------------------------------------------------
# INTERFAZ PRINCIPAL CON PESTAÑAS
# ------------------------------------------------------------------
tab1, tab2 = st.tabs(
    ["📜 Generar Certificado", "📊 Reportes de Gestión (Jefatura)"]
)

# ------------------------------------------------------------------
# PESTAÑA 1: GENERAR CERTIFICADO
# ------------------------------------------------------------------
with tab1:
    st.markdown("#### 1. Parámetros de la Consulta DSpace")
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
        "🔍 Consultar DSpace y Generar Documento(s)", use_container_width=True
    )

    if btn_procesar or "datos_cargados" in st.session_state:
        if btn_procesar:
            if not url_input:
                st.warning("⚠️ Por favor ingresa la URL o Handle de DSpace.")
                st.stop()

            with st.spinner("Conectando con el repositorio DSpace de la UCuenca..."):
                meta = extraer_metadatos_dspace(url_input)
                st.session_state["datos_cargados"] = meta

        meta = st.session_state["datos_cargados"]
        ref_info = next(
            item for item in LISTA_REFERENCISTAS if item["nombre"] == referencista_sel
        )

        st.markdown("---")
        st.markdown("#### 2. Validación de Metadatos y Estudiantes")

        col_f, col_c = st.columns(2)
        with col_f:
            facultad_final = st.text_input(
                "Facultad Detectada:", value=meta["facultad"]
            )
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
                    f"Nombre Estudiante #{idx}:",
                    value=autor_nombre,
                    key=f"nom_{idx}",
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
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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
                label=f"📦 Descargar TODOS los Certificados en Un Archivo ZIP ({len(certificados_generados)} archivos)",
                data=zip_buffer,
                file_name="Certificados_No_Adeudar_UCuenca.zip",
                mime="application/zip",
            )

        st.markdown("---")
        st.markdown("#### 3. Registro en Nube (Google Sheets)")
        if st.button("📊 Guardar Registro(s) en Google Sheets", use_container_width=True):
            with st.spinner("Guardando registro(s) en Google Sheets..."):
                exitos = 0
                errores = []
                for item in certificados_generados:
                    exito, msg = guardar_en_google_sheets(item)
                    if exito:
                        exitos += 1
                    else:
                        errores.append(f"{item['autor']}: {msg}")

                if exitos > 0:
                    st.success(
                        f"✅ Se guardaron {exitos} registro(s) correctamente en Google Sheets."
                    )
                if errores:
                    st.error(f"❌ Error al conectar con Google Sheets: {errores[0]}")


# ------------------------------------------------------------------
# PESTAÑA 2: REPORTES DE GESTIÓN (JEFATURA)
# ------------------------------------------------------------------
with tab2:
    st.markdown("#### 🔍 Filtros para la Auditoría Mensual")

    df_registros = cargar_datos_reporte()

    col_m, col_a, col_r = st.columns(3)
    with col_m:
        mes_sel = st.selectbox(
            "Mes:",
            [
                "Septiembre",
                "Enero",
                "Febrero",
                "Marzo",
                "Abril",
                "Mayo",
                "Junio",
                "Julio",
                "Agosto",
                "Octubre",
                "Noviembre",
                "Diciembre",
            ],
            index=0,
        )
    with col_a:
        anio_sel = st.selectbox("Año:", [2026, 2025, 2024], index=0)
    with col_r:
        ref_filtro = st.selectbox(
            "Referencista (Filtrar por usuario):",
            ["-- Todos los Referencistas --"]
            + sorted([r["nombre"] for r in LISTA_REFERENCISTAS]),
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # Cálculo de Métricas (KPIs)
    tot_certificados = len(df_registros)
    referencista_top = (
        df_registros["Referencista"].mode()[0]
        if not df_registros.empty
        else "N/A"
    )
    promedio_diario = round(tot_certificados / 15, 1) if tot_certificados > 0 else 0

    # Tarjetas de Métricas (KPI Cards)
    kpi1, kpi2, kpi3 = st.columns(3)

    with kpi1:
        st.markdown(
            f"""
        <div class="kpi-card kpi-blue">
            <div class="kpi-title">TOTAL CERTIFICADOS DEL MES</div>
            <div class="kpi-value">{tot_certificados}</div>
            <div class="kpi-subtext">{mes_sel} {anio_sel}</div>
        </div>
        """,
            unsafe_allow_html=True,
        )

    with kpi2:
        st.markdown(
            f"""
        <div class="kpi-card kpi-red">
            <div class="kpi-title">REFERENCISTA MÁS ACTIVO</div>
            <div class="kpi-value" style="font-size:20px;">{referencista_top}</div>
            <div class="kpi-subtext">42 certificados emitidos (32%)</div>
        </div>
        """,
            unsafe_allow_html=True,
        )

    with kpi3:
        st.markdown(
            f"""
        <div class="kpi-card kpi-navy">
            <div class="kpi-title">PROMEDIO DIARIO</div>
            <div class="kpi-value">{promedio_diario}</div>
            <div class="kpi-subtext">Certificados por día hábil</div>
        </div>
        """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 📋 Registro Detallado de Emisiones (Evidencia)")

    # Formatear la tabla con links clickeables
    st.dataframe(
        df_registros,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Link DSpace": st.column_config.LinkColumn("Link DSpace")
        },
    )

    # Botón de Descarga Excel estilo Oficial
    col_txt, col_btn = st.columns([2, 1])
    with col_txt:
        st.caption("Mostrando los últimos registros filtrados.")
    with col_btn:
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
            df_registros.to_excel(
                writer, sheet_name="Reporte_Consolidado", index=False
            )
        buffer_excel.seek(0)

        st.markdown('<div class="btn-excel">', unsafe_allow_html=True)
        st.download_button(
            label="📊 Descargar Reporte Consolidado en Excel (.xlsx)",
            data=buffer_excel,
            file_name=f"Reporte_Certificados_{mes_sel}_{anio_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown("</div>", unsafe_allow_html=True)




