import streamlit as st
import zipfile
import io
from docx import Document
from fpdf import FPDF

# Configuración de la página
st.set_page_config(page_title="Convertidor", page_icon="📄")

def procesar_archivos():
    st.title("📂 Convertidor Word a PDF")
    st.write("Sube tu carpeta o archivos .docx")

    archivos = st.file_uploader("Selecciona archivos", accept_multiple_files=True)

    if archivos:
        # Filtrar solo archivos .docx validos
        validos = [f for f in archivos if f.name.lower().endswith(".docx") and not f.name.startswith("~$")]
        
        if not validos:
            st.warning("No se encontraron archivos .docx.")
            return

        st.info(f"Archivos a convertir: {len(validos)}")

        if st.button("🚀 Iniciar conversión"):
            zip_buffer = io.BytesIO()
            progreso = st.progress(0)
            
            with zipfile.ZipFile(zip_buffer, "w") as z:
                for i, arc in enumerate(validos):
                    try:
                        # Leer Word
                        doc = Document(arc)
                        pdf = FPDF()
                        pdf.add_page()
                        pdf.set_font("Helvetica", size=12)
                        
                        # Escribir párrafos
                        for p in doc.paragraphs:
                            if p.text.strip():
                                # Limpiar texto para evitar errores de símbolos
                                limpio = p.text.encode('latin-1', 'replace').decode('latin-1')
                                pdf.multi_cell(0, 10, txt=limpio)
                        
                        # Generar PDF
                        pdf_output = pdf.output()
                        nombre_pdf = arc.name.rsplit('.', 1)[0] + ".pdf"
                        z.writestr(nombre_pdf, pdf_output)
                        
                        progreso.progress((i + 1) / len(validos))
                    except Exception as e:
                        st.error(f"Error en {arc.name}: {str(e)}")

            st.success("¡Conversión completa!")
            st.download_button(
                label="⬇️ Descargar ZIP",
                data=zip_buffer.getvalue(),
                file_name="documentos_pdf.zip",
                mime="application/zip"
            )

if __name__ == "__main__":
    procesar_archivos()
