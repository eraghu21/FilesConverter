import streamlit as st
from io import BytesIO
from pdf2docx import Converter
from docx2pdf import convert
from pdf2image import convert_from_bytes
from PIL import Image
import fitz  # PyMuPDF
import os
import tempfile
import zipfile
import platform
from fpdf import FPDF

st.set_page_config(page_title="Universal File Converter", page_icon="📄", layout="centered")

st.title("📄 Universal File Converter")
st.markdown("Convert between **Word, PDF, and Image** formats — preserving full content, layout, and quality!")

uploaded_file = st.file_uploader("📤 Upload your file", type=["pdf", "docx", "jpg", "jpeg", "png"])

conversion_type = st.selectbox(
    "🔄 Choose conversion type",
    ["Select", "Word ➜ PDF", "PDF ➜ Word", "PDF ➜ Image(s)", "Image ➜ PDF"]
)

if uploaded_file and conversion_type != "Select":
    if st.button("🚀 Convert"):
        try:
            temp_dir = tempfile.mkdtemp()

            # Save uploaded file temporarily
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.read())

            # -------------------------------
            # Word ➜ PDF
            # -------------------------------
            if conversion_type == "Word ➜ PDF":
                output_path = os.path.join(temp_dir, "converted.pdf")
                try:
                    # Use native docx2pdf for full layout conversion
                    convert(input_path, output_path)
                except Exception:
                    # Fallback method using FPDF if Word not available
                    from docx import Document
                    doc = Document(input_path)
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", size=12)
                    for para in doc.paragraphs:
                        text = para.text.encode("latin-1", "replace").decode("latin-1")
                        pdf.multi_cell(0, 10, txt=text)
                    pdf.output(output_path)

                with open(output_path, "rb") as f:
                    st.download_button("⬇️ Download PDF",
                                       data=f.read(),
                                       file_name="converted.pdf",
                                       mime="application/pdf")

            # -------------------------------
            # PDF ➜ Word
            # -------------------------------
            elif conversion_type == "PDF ➜ Word":
                output_path = os.path.join(temp_dir, "converted.docx")
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()

                with open(output_path, "rb") as f:
                    st.download_button("⬇️ Download Word",
                                       data=f.read(),
                                       file_name="converted.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            # -------------------------------
            # PDF ➜ Image(s)
            # -------------------------------
            elif conversion_type == "PDF ➜ Image(s)":
                pdf_document = fitz.open(input_path)
                zip_buffer = BytesIO()

                with zipfile.ZipFile(zip_buffer, "a") as zipf:
                    for page_num in range(len(pdf_document)):
                        page = pdf_document.load_page(page_num)
                        pix = page.get_pixmap(dpi=200)
                        img_bytes = BytesIO(pix.tobytes("png"))
                        zipf.writestr(f"page_{page_num+1}.png", img_bytes.getvalue())

                zip_buffer.seek(0)
                st.image(convert_from_bytes(open(input_path, "rb").read(), dpi=100)[0], caption="Preview Page 1")
                st.download_button("⬇️ Download All Pages (ZIP)", data=zip_buffer,
                                   file_name="pdf_images.zip", mime="application/zip")

            # -------------------------------
            # Image ➜ PDF
            # -------------------------------
            elif conversion_type == "Image ➜ PDF":
                image = Image.open(input_path)
                pdf_buffer = BytesIO()
                image.convert("RGB").save(pdf_buffer, format="PDF", resolution=100.0)
                st.download_button("⬇️ Download PDF", data=pdf_buffer.getvalue(),
                                   file_name="converted.pdf", mime="application/pdf")

        except Exception as e:
            st.error(f"❌ Conversion failed: {e}")
