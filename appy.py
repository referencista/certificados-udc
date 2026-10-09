import io
import re
import unicodedata
import zipfile
from datetime import datetime, timedelta, timezone

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
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
    page_title="UCuenca - Generador de Certificados de No Adeudar",
    page_icon="🎓",
    layout="wide",
)

CSS_UCUENCA = """
<style>
    /* Estilo General */
    body {
        font-family: 'Arial', sans-serif;
        background-color: #f4f6f9;
    }
    
    /* 1. Header Institucional (Azul con borde inferior rojo) */
    .uc-header {
        text-align: center;
        background-color: #13386c !important;
        color: white;
        padding: 20px 30px;
        border-radius: 8px 8px 6px 6px;
        margin-bottom: 25px;
        border-bottom: 5px solid #b81c32 !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.08);
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

    /* 2. Botones (Azul con franja roja inferior) */
    .stButton>button, .stDownloadButton>button {
        background-color: #13386c !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: none !important;
        border-bottom: 4px solid #b81c32 !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease !important;
    }
    
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #0b2347 !important;
        border-bottom: 4px solid #d4223b !important;
        color: #ffffff !important;
    }

    /* 3. ELIMINAR BARRA SUPERIOR */
    header[data-testid="stHeader"],
    div[data-testid="stToolbar"],
    div[data-testid="stDecoration"] {
        display: none !important;
    }

    /* 4. ELIMINAR ÍCONOS FLOTANTES INFERIORES */
    footer,
    [data-testid="manage-app-button"],
    button[title="Manage app"],
    div[class*="stAppDeployButton"],
    div[data-testid="stStatusWidget"],
    div[class*="viewerBadge"],
    .viewerBadge_container__1QSob {
        display: none !important;
    }
</style>
"""

# Inyección en Streamlit
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
# LISTA OFICIAL DE REFERENCISTAS Y ESTRUCTURAS
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
    {"nombre": "JOANNA NOEMI MOGOLLÓN GUZMÁN", "cargo": "Bibliotecario 2"},
    {"nombre": "PAOLA DEL ROCIO AMAYA ARCE", "cargo": "Bibliotecario 2"},
    {"nombre": "PATRICIA MARIBEL DUCHI PESÁNTEZ", "cargo": "Bibliotecario 2"},
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
    texto = unicodedata.normalize("NFD", str(texto))
    texto = re.sub(r"[\u0300-\u036f]", "", texto)
    return texto.lower()


# ==================================================================
# SISTEMA DE FORMATEO DE CARRERAS (CATÁLOGO CANÓNICO + FALLBACK)
# ==================================================================

CATALOGO_OFICIAL_CARRERAS = {
    # Pregrado
    "pedagogia de las ciencias experimentales": "Pedagogía de las Ciencias Experimentales",
    "pedagogia de la actividad fisica y deporte": "Pedagogía de la Actividad Física y Deporte",
    "pedagogia de los idiomas nacionales y extranjeros": "Pedagogía de los Idiomas Nacionales y Extranjeros",
    "pedagogia de las artes y las humanidades": "Pedagogía de las Artes y las Humanidades",
    "educacion basica": "Educación Básica",
    "educacion inicial": "Educación Inicial",
    "comunicacion": "Comunicación",
    "sociologia": "Sociología",
    "trabajo social": "Trabajo Social",
    "derecho": "Derecho",
    "genero y desarrollo": "Género y Desarrollo",
    "orientacion familiar": "Orientación Familiar",
    "psicologia": "Psicología",
    "psicologia clinica": "Psicología Clínica",
    "psicologia social": "Psicología Social",
    "contabilidad y auditoria": "Contabilidad y Auditoría",
    "administracion de empresas": "Administración de Empresas",
    "economia": "Economía",
    "mercadotecnia": "Mercadotecnia",
    "finanzas": "Finanzas",
    "ingenieria civil": "Ingeniería Civil",
    "ingenieria de sistemas": "Ingeniería de Sistemas",
    "electronica y telecomunicaciones": "Electrónica y Telecomunicaciones",
    "ingenieria electrica": "Ingeniería Eléctrica",
    "ingenieria industrial": "Ingeniería Industrial",
    "ingenieria quimica": "Ingeniería Química",
    "ingenieria ambiental": "Ingeniería Ambiental",
    "bioquimica y farmacia": "Bioquímica y Farmacia",
    "medicina": "Medicina",
    "enfermeria": "Enfermería",
    "fisioterapia": "Fisioterapia",
    "nutricion y dietetica": "Nutrición y Dietética",
    "fonoaudiologia": "Fonoaudiología",
    "imagenologia y radiologia": "Imagenología y Radiología",
    "medicina veterinaria": "Medicina Veterinaria",
    "agronomia": "Agronomía",
    "arquitectura": "Arquitectura",
    "diseno grafico": "Diseño Gráfico",
    "diseno de interiores": "Diseño de Interiores",
    "artes visuales": "Artes Visuales",
    "gastronomia": "Gastronomía",
    "turismo": "Turismo",
    # Posgrado / Maestrías
    "maestria en contabilidad y auditoria": "Maestría en Contabilidad y Auditoría",
    "maestria en gestion publica y buen gobierno": "Maestría en Gestión Pública y Buen Gobierno",
    "maestria en educacion": "Maestría en Educación",
    "maestria en psicologia": "Maestría en Psicología",
}

PALABRAS_TILDADAS_FALLBACK = {
    "pedagogia": "pedagogía",
    "ingenieria": "ingeniería",
    "psicologia": "psicología",
    "sociologia": "sociología",
    "agronomia": "agronomía",
    "gastronomia": "gastronomía",
    "hoteleria": "hotelería",
    "fonoaudiologia": "fonoaudiología",
    "imagenologia": "imagenología",
    "auditoria": "auditoría",
    "maestria": "maestría",
    "biologia": "biología",
    "tecnologia": "tecnología",
    "educacion": "educación",
    "comunicacion": "comunicación",
    "administracion": "administración",
    "gestion": "gestión",
    "investigacion": "investigación",
    "orientacion": "orientación",
    "nutricion": "nutrición",
    "basica": "básica",
    "clinica": "clínica",
    "medica": "médica",
    "fisica": "física",
    "quimica": "química",
    "bioquimica": "bioquímica",
    "politica": "política",
    "publica": "pública",
    "economica": "económica",
    "electronica": "electrónica",
    "electrica": "eléctrica",
    "musica": "música",
    "genero": "género",
    "diseno": "diseño",
    "rediseno": "rediseño",
}

CONECTORES_ESPANOL = {
    "de", "del", "la", "las", "los", "el", "y", "e", "o", "u", "en", "con", "por", "para", "a"
}


def formatear_carrera_espanol(texto):
    if not texto:
        return ""

    texto_limpio = re.sub(r"\s+", " ", str(texto)).strip()
    texto_norm = normalizar_texto(texto_limpio)
