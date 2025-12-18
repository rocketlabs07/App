import streamlit as st
import zipfile
import io
from docx import Document
from fpdf import FPDF

st.set_page_config(page_title="Conversor de Carpetas", page_icon="📂")

st.title("📂 Conversor de Word a PDF")
st.write("Selecciona la carpeta o todos los archivos. Solo procesaremos los .docx")

# Cargador de archivos
archivos_subidos = st.file_uploader(
    "Sube tus archivos aquí", 
    accept_multiple_files=True
)

if archivos_subidos:
    # Filtramos archivos válidos
    docs_a_procesar = [f for f in archivos_subidos if f.name.lower().endswith(".docx") and not f.name.startswith("~$")]
    
    if not docs_a_procesar:
        st.warning("No se encontraron archivos .docx válidos.")
    else:
        st.info(f"Archivos detectados: {len(docs_a_procesar)}")

        if st.button("🚀 Convertir a PDF"):
            buf = io.BytesIO()
            barra = st.progress(0)
            
            try:
                with zipfile.ZipFile(buf, "w") as z:
                    for i, archivo in enumerate(docs_a_procesar):
                        # Leer Word
                        doc = Document(archivo)
                        pdf = FPDF()
                        pdf.add_page()
                        pdf.set_font("Arial", size=12)
                        
                        # Escribir contenido
                        for para in doc.paragraphs:
                            if para.text.strip():
                                # fpdf2 maneja mejor el texto, pero usamos 'latin-1' por compatibilidad simple
                                txt = para.text.encode('latin-1', 'replace').decode('latin-1')
                                pdf.multi_cell(0, 10, txt=txt)
                        
                        # Obtener bytes del PDF
                        pdf_bytes = pdf.output()
                        
                        # Nombre de salida
                        nombre_pdf = archivo.name.rsplit('.', 1)[0] + ".pdf"
                        z.writestr(nombre_pdf, pdf_bytes)
                        
                        barra.progress((i + 1) / len(docs_a_procesar))

                st.success("¡Hecho!")
                st.download_button(
                    label="⬇️ Descargar ZIP",
                    data=buf.getvalue(),
                    file_name="mis_pdfs.zip",
                    mime="application/zip"
                )
            except Exception as e:
                st.error(f"Ocurrió un error técnico: {e}")
