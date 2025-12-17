import streamlit as st
import os
import zipfile
from docx import Document
from fpdf import FPDF
import io

st.set_page_config(page_title="Word a PDF - Express", page_icon="📄")

st.title("📂 Conversor de Word a PDF")
st.info("Sube tus archivos .docx y los convertiremos a PDF en un solo paso.")

archivos_subidos = st.file_uploader("Arrastra aquí tus archivos Word", type=["docx"], accept_multiple_files=True)

if archivos_subidos:
    # Creamos un archivo ZIP en memoria para que sea rápido
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for archivo in archivos_subidos:
            # Leer el contenido del Word
            doc = Document(archivo)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            
            for para in doc.paragraphs:
                # Añadir cada párrafo al PDF (limpiando caracteres especiales)
                pdf.multi_cell(0, 10, txt=para.text.encode('latin-1', 'replace').decode('latin-1'))
            
            # Guardar PDF en memoria y añadir al ZIP
            pdf_output = pdf.output(dest='S').encode('latin-1')
            nombre_pdf = archivo.name.replace(".docx", ".pdf")
            z.writestr(nombre_pdf, pdf_output)

    st.success(f"¡{len(archivos_subidos)} archivos procesados!")
    
    st.download_button(
        label="⬇️ Descargar todos los PDFs (.zip)",
        data=buf.getvalue(),
        file_name="archivos_convertidos.zip",
        mime="application/zip"
    )


