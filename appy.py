import io
import re
from datetime import datetime
from bs4 import BeautifulSoup
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import requests
import streamlit as st

# Configuración de la página web
st.set_page_config(
    page_title="Generador de Certificados - UCuenca", page_icon="📜"
)

st.title("📜 Generador de Certificado de No Adeudar")
st.subheader("Centro de Documentación Regional 'Juan Bautista Vázquez'")


# ---------------------------------------------------------
# FUNCION PARA EXTRAER METADATOS DE DSPACE 7
# ---------------------------------------------------------
def obtener_metadatos_dspace(url):
  headers = {
      "User-Agent": (
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      )
  }
  try:
    response = requests.get(url, headers=headers, timeout=10)
    if response.status_code != 200:
      return None, "No se pudo acceder a la URL proporcionada."

    soup = BeautifulSoup(response.text, "html.parser")

    # Extraer Autor
    autor_elem = soup.find("meta", {"name": "DC.creator"}) or soup.find(
        "meta", {"name": "citation_author"}
    )
    autor = autor_elem["content"].strip() if autor_elem else ""

    # Extraer Handle/URI permanente
    handle_elem = soup.find("meta", {"name": "DC.identifier.uri"})
    handle_link = (
        handle_elem["content"].strip() if handle_elem else url
    )  # Si no hay handle, usa la URL ingresada

    # Extraer Facultad / Publisher
    publisher_elem = soup.find("meta", {"name": "DC.publisher"})
    facultad = publisher_elem["content"].strip() if publisher_elem else ""

    # Extraer Titulación / Carrera si existe en los metadatos
    subject_elem = soup.find("meta", {"name": "DC.subject"})
    carrera = subject_elem["content"].strip() if subject_elem else ""

    return {
        "autor": autor.upper(),
        "facultad": facultad,
        "carrera": carrera,
        "handle": handle_link,
    }, None
  except Exception as e:
    return None, f"Error al procesar la página: {str(e)}"


# ---------------------------------------------------------
# LISTA DE REFERENCISTAS / BIBLIOTECARIOS
# ---------------------------------------------------------
# Puedes editar o agregar más personas a esta lista
LISTA_REFERENCISTAS = [
    {
        "nombre": "PAOLA DEL ROCÍO AMAYA ARCE",
        "cargo": "Bibliotecario 2",
    },
    {
        "nombre": "NARCISA DE JESÚS CÁRDENAS",
        "cargo": "Referencista / Bibliotecario",
    },
    {
        "nombre": "JUAN CARLOS PÉREZ",
        "cargo": "Bibliotecario 1",
    },
]

# ---------------------------------------------------------
# FORMULARIO EN LA INTERFAZ WEB
# ---------------------------------------------------------
st.markdown("---")

col1, col2 = st.columns([2, 1])

with col1:
  url_input = st.text_input(
      "URL del trabajo de titulación en DSpace:",
      placeholder="https://dspace.ucuenca.edu.ec/items/...",
  )

with col2:
  cedula_input = st.text_input(
      "Número de Cédula del estudiante:", placeholder="1401064256"
  )

col3, col4 = st.columns(2)

with col3:
  nivel_academico = st.selectbox(
      "Nivel Académico:", ["Pregrado", "Maestría / Posgrado"]
  )

with col4:
  # Selección del Referencista
  nombres_ref = [r["nombre"] for r in LISTA_REFERENCISTAS]
  referencista_sel = st.selectbox(
      "Bibliotecario / Referencista que emite:", nombres_ref
  )

# Obtener cargo del referencista seleccionado
cargo_sel = next(
    r["cargo"] for r in LISTA_REFERENCISTAS if r["nombre"] == referencista_sel
)

# Configurar fecha por defecto (Fecha actual)
meses = [
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
hoy = datetime.now()
fecha_defecto = f"{hoy.day} de {meses[hoy.month-1]} de {hoy.year}"
fecha_str = st.text_input("Fecha del certificado:", value=fecha_defecto)

# ---------------------------------------------------------
# BOTÓN PARA GENERAR EL DOCUMENTO WORD (.DOCX)
# ---------------------------------------------------------
if st.button("📄 Generar Certificado (.docx)", type="primary"):
  if not url_input or not cedula_input:
    st.error(
        "Por favor, ingresa tanto la URL de DSpace como el número de cédula."
    )
  else:
    with st.spinner("Consultando DSpace y generando documento..."):
      datos, error = obtener_metadatos_dspace(url_input)

      if error:
        st.error(error)
      else:
        # Formatear el documento Word
        doc = docx.Document()

        # Encabezado estándar UC-CDRJVB-FOR-020
        p_head1 = doc.add_paragraph()
        p_head1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p_head1.add_run(
            "FORMATO DE NO ADEUDAR MATERIAL BIBLIOGRÁFICO A LA BIBLIOTECA\n"
        )
        r1.bold = True
        r1.font.size = Pt(10)

        r2 = p_head1.add_run("UC-CDRJVB-FOR-020\nPágina 1 de 1")
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(100, 100, 100)

        doc.add_paragraph("\n")

        # Título principal
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_title = p_title.add_run("CERTIFICADO DE NO ADEUDAR")
        r_title.bold = True
        r_title.font.size = Pt(14)

        doc.add_paragraph("\n")

        # Ajuste según Pregrado o Maestría
        facultad_txt = (
            datos["facultad"]
            if datos["facultad"]
            else "[Nombre de la Facultad]"
        )
        carrera_txt = (
            datos["carrera"] if datos["carrera"] else "[Carrera / Programa]"
        )

        if nivel_academico == "Pregrado":
          estudio_txt = f"estudiante de la {facultad_txt} de la Carrera de {carrera_txt}"
        else:
          estudio_txt = f"estudiante del Programa de {carrera_txt} de la {facultad_txt}"

        # Cuerpo del certificado
        p_body = doc.add_paragraph()
        p_body.paragraph_format.line_spacing = 1.15
        p_body.paragraph_format.space_after = Pt(12)

        p_body.add_run(
            'El Centro de Documentación Regional "Juan Bautista Vázquez"'
            " certifica que "
        )
        r_est = p_body.add_run(f"{datos['autor']}")
        r_est.bold = True

        p_body.add_run(", portador de la cédula de ciudadanía No. ")
        r_ced = p_body.add_run(f"{cedula_input}")
        r_ced.bold = True

        p_body.add_run(
            f", {estudio_txt}, no adeuda ningún bien, ni material"
            " bibliográfico en esta dependencia."
        )

        doc.add_paragraph("\n")

        # Fecha
        p_date = doc.add_paragraph(f"Cuenca, {fecha_str}")
        p_date.paragraph_format.space_after = Pt(24)

        # Atentamente y Firma
        p_sign = doc.add_paragraph("Atentamente,\n\n\n")
        p_sign.add_run(
            "________________________________________\n"
        ).bold = True

        r_ref = p_sign.add_run(f"{referencista_sel}\n")
        r_ref.bold = True

        p_sign.add_run(f"{cargo_sel}\n")
        p_sign.add_run('CDR "Juan Bautista Vázquez"\n\n')

        # Link Handle
        r_link = p_sign.add_run(f"Link: {datos['handle']}")
        r_link.font.size = Pt(9)
        r_link.font.italic = True

        # Guardar en memoria para descarga
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("¡Certificado generado correctamente!")

        st.download_button(
            label="📥 Descargar Certificado en Word (.docx)",
            data=buffer,
            file_name=(
                f"Certificado_No_Adeudar_{cedula_input if cedula_input else 'UC'}.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )
