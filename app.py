import streamlit as st
import zipfile
import io
from docx import Document
from fpdf import FPDF

st.set_page_config(page_title="Conversor Masivo de Word", page_icon="📁")

st.title("📁 Conversor de Carpetas Word a PDF")
st.write("Selecciona todos los archivos de tu carpeta (puedes arrastrar todo el contenido). Solo se convertirán los archivos .docx.")

# El cargador de archivos ahora acepta múltiples archivos
# Nota: El usuario puede seleccionar todos los archivos dentro de la carpeta (Cmd+A en Mac) 
# y arrastrarlos aquí.
archivos_subidos = st.file_uploader(
    "Arrastra todos los archivos y subcarpetas aquí", 
    type=["docx"], 
    accept_multiple_files=True
)

if archivos_subidos:
    # Filtrar solo los archivos que terminan en .docx y no son temporales
    docs_a_procesar = [f for f in archivos_subidos if f.name.endswith(".docx") and not f.name.startswith("~$")]
    
    if len(docs_a_procesar) == 0:
        st.warning("No se encontraron archivos .docx válidos en la selección.")
    else:
        st.info(f"Se han detectado {len(docs_a_procesar)} archivos Word para convertir.")

        if st.button("🚀 Iniciar Conversión Masiva"):
            buf = io.BytesIO()
            progreso = st.progress(0)
            
            with zipfile.ZipFile(buf, "w") as z:
                for i, archivo in enumerate(docs_a_procesar):
                    try:
                        # Leer el Word
                        doc = Document(archivo)
                        pdf = FPDF()
                        pdf.add_page()
                        # Usar una fuente estándar que soporte mejor caracteres
                        pdf.set_font("Helvetica", size=12)
                        
                        for para in doc.paragraphs:
                            if para.text.strip():
                                # Limpieza básica para evitar errores de codificación en PDF estándar
                                texto_limpio = para.text.encode('latin-1', 'replace').decode('latin-1')
                                pdf.multi_cell(0, 10, txt=texto_limpio)
                        
                        # Guardar PDF en el ZIP
                        pdf_output = pdf.output(dest='S').encode('latin-1')
                        nombre_pdf = archivo.name.replace(".docx", ".pdf")
                        z.writestr(nombre_pdf, pdf_output)
                        
                    except Exception as e:
                        st.error(f"Error en {archivo.name}: {e}")
                    
                    # Actualizar barra de progreso
                    progreso.progress((i + 1) / len(docs_a_procesar))

            st.success("¡Conversión finalizada!")
            
            st.download_button(
                label="⬇️ Descargar Carpeta de PDFs (.zip)",
                data=buf.getvalue(),
                file_name="todos_los_pdfs.zip",
                mime="application/zip"
            )
