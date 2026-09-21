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
# EXTRACCIÓN AVANZADA DE METADATOS VÍA API REST DSPACE 7 (UCUENCA)
# ------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def extraer_metadatos_dspace(url_input):
    url_clean = url_input.strip()

    # Detectar Handle (ej: 123456789/48627) o UUID (ej: ad19c5fb-5289-4881-89fa-8a1e831345c6)
    match_handle = re.search(r'(\d+/\d+)', url_clean)
    match_uuid = re.search(r'([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})', url_clean, re.I)

    base_apis = [
        "https://rest-dspace.ucuenca.edu.ec/server/api",
        "https://dspace.ucuenca.edu.ec/server/api"
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

    headers = {'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0'}
    data = None
    
    if endpoint:
        for base in base_apis:
            try:
                resp = requests.get(base + endpoint, headers=headers, timeout=10, verify=False)
                if resp.status_code == 200:
                    data = resp.json()
                    break
            except Exception:
                continue

    # Si la API no responde
    if not data:
        return {
            "autores": ["APELLIDOS, NOMBRES ESTUDIANTE"],
            "facultad": "",
            "carrera": "",
            "handle": handle_official
        }

    metadata = data.get("metadata", {})

    def get_meta_value(keys_list, default=""):
        """Busca el valor probando múltiples claves posibles en DSpace 7."""
        if isinstance(keys_list, str):
            keys_list = [keys_list]
        for key in keys_list:
            items = metadata.get(key, [])
            if items and len(items) > 0:
                val = items[0].get("value", "").strip()
                if val:
                    return val
        return default

    # 1. EXTRAER AUTORES EN MAYÚSCULAS CONSERVANDO LAS COMAS Y EL ORDEN ORIGINAL (APELLIDOS, NOMBRES)
    raw_authors = [a.get("value") for a in metadata.get("dc.contributor.author", []) if a.get("value")]
    autores_formateados = []
    for a in raw_authors:
        # Limpia espacios extra, conserva la coma y convierte a MAYÚSCULAS
        a_clean = re.sub(r'\s+', ' ', a).strip().upper()
        if a_clean and a_clean not in autores_formateados:
            autores_formateados.append(a_clean)

    # 2. EXTRAER FACULTAD Y CARRERA REVISANDO MÚLTIPLES ETIQUETAS POSIBLES
    facultad = get_meta_value([
        "thesis.degree.grantor",
        "dc.publisher",
        "dc.department",
        "dc.contributor.department"
    ], default="")

    carrera = get_meta_value([
        "thesis.degree.discipline",
        "thesis.degree.name",
        "dc.degree.discipline",
        "dc.subject"
    ], default="")

    # 3. EXTRAER EL HANDLE OFICIAL SI LA ENTRADA FUE UN UUID
    uri_items = metadata.get("dc.identifier.uri", [])
    for uri in uri_items:
        val_uri = uri.get("value", "")
        if "handle/" in val_uri:
            match_h = re.search(r'(\d+/\d+)', val_uri)
            if match_h:
                handle_official = f"https://dspace.ucuenca.edu.ec/handle/{match_h.group(1)}"
                break

    return {
        "autores": autores_formateados if autores_formateados else ["APELLIDOS, NOMBRES ESTUDIANTE"],
        "facultad": facultad,
        "carrera": carrera,
        "handle": handle_official
    }

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
