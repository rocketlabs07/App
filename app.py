import streamlit as st
import zipfile
import io
from docx import Document
from fpdf import FPDF

st.set_page_config(page_title="Conversor de Carpetas", page_icon="📂")

st.title("📂 Conversor Automático de Carpetas")
st.markdown("""
1. Haz clic en **Browse files**.
2. Selecciona la **carpeta principal** (la que contiene todo).
3. El programa buscará todos los Word dentro de ella y sus subcarpetas.
""")

# Cambiamos la configuración para aceptar carpetas si el navegador lo permite
# Aunque Streamlit lo muestra como archivos, al arrastrar una carpeta, 
# el navegador procesa todos los archivos internos.
archivos_subidos = st.file_uploader(
    "Sube tu carpeta completa aquí", 
    accept_multiple_files=True
)

if archivos_subidos:
    # Filtrado inteligente
    docs_a_procesar = [f for f in archivos_subidos if f.name.lower().endswith(".docx") and not f.name.startswith("~$")]
    
    if not docs_a_procesar:
        st.warning("No se encontraron archivos .docx en la carpeta seleccionada.")
    else:
        st.success(f"Se encontraron {len(docs_a_procesar)} archivos Word. Los demás formatos han sido descartados.")

        if st.button("🚀 Convertir todo a PDF"):
            buf = io.BytesIO()
            barra = st.progress(0)
            
            with zipfile.ZipFile(buf, "w") as z:
                for i, archivo in enumerate(docs_a_procesar):
                    try:
                        # Procesamiento del Word
                        doc = Document(archivo)
                        pdf = FPDF()
                        pdf.add_page()
                        pdf.set_font("Arial", size=12)
                        
                        for para in doc.paragraphs:
                            if para.text.strip():
                                # Manejo de caracteres latinos
                                txt = para.text.encode('latin-1', 'replace').decode('latin-1')
                                pdf.multi_cell(0, 10, txt=txt)
                        
                        pdf_bytes = pdf.output(dest='S').encode('latin-1')
                        
                        # Guardamos en el ZIP usando el nombre original
                        nombre_pdf = archivo.name.rsplit('.', 1)[0] + ".pdf"
                        z.writestr(nombre_pdf, pdf_bytes)
                        
                    except Exception as e:
                        st.error(f"Error procesando {archivo.name}: {e}")
                    
                    barra.progress((i + 1) / len(docs_a_procesar))

            st.download_button(
                label="⬇️ Descargar todos los PDFs (.zip)",
                data=buf.getvalue(),
                file_name="conversion_carpeta.zip",
                mime="application/zip"
            )
