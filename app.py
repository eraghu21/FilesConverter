import streamlit as st
from io import BytesIO
from fpdf import FPDF
from PIL import Image
from pdf2image import convert_from_bytes
from docx import Document
import fitz  # PyMuPDF
import pdfplumber
import zipfile
import os

st.set_page_config(page_title="Universal File Converter", page_icon="📄", layout="centered")

st.title("📄 Universal File Converter")
st.markdown("Convert between **Word, PDF, and Image** formats with full accuracy and quality.")

uploaded_file = st.file_uploader("📤 Upload a file", type=["pdf", "docx", "jpg", "jpeg", "png"])

conversion_type = st.selectbox(
    "🔄 Choose conversion type",
    ["Select", "Word ➜ PDF", "PDF ➜ Word", "PDF ➜ Image(s)", "Image ➜ PDF"]
)

if uploaded_file and conversion_type != "Select":
    if st.button("🚀 Convert"):
        try:
            if conversion_type == "Word ➜ PDF":
                # Convert Word → PDF (preserve text line breaks)
                doc = Document(uploaded_file)
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                for para in doc.paragraphs:
                    text = para.text.encode("latin-1", "replace").decode("latin-1")
                    pdf.multi_cell(0, 10, txt=text)
                pdf_buffer = BytesIO()
                pdf.output(pdf_buffer)
                pdf_buffer.seek(0)
                st.download_button("⬇️ Download PDF", data=pdf_buffer, file_name="converted.pdf", mime="application/pdf")

            elif conversion_type == "PDF ➜ Word":
                # Convert PDF → Word (text layout preserved using pdfplumber)
                with pdfplumber.open(uploaded_file) as pdf:
                    doc = Document()
                    for page in pdf.pages:
                        text = page.extract_text() or ""
                        doc.add_paragraph(text)
                doc_buffer = BytesIO()
                doc.save(doc_buffer)
                st.download_button("⬇️ Download Word", data=doc_buffer.getvalue(),
                                   file_name="converted.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            elif conversion_type == "PDF ➜ Image(s)":
                # Convert PDF → multiple images (ZIP download)
                pages = convert_from_bytes(uploaded_file.read(), dpi=300)
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a") as zipf:
                    for i, page in enumerate(pages):
                        img_bytes = BytesIO()
                        page.save(img_bytes, format="PNG")
                        zipf.writestr(f"page_{i+1}.png", img_bytes.getvalue())
                zip_buffer.seek(0)
                st.image(pages[0], caption="Preview: Page 1")
                st.download_button("⬇️ Download All Pages (ZIP)", data=zip_buffer,
                                   file_name="pdf_images.zip", mime="application/zip")

            elif conversion_type == "Image ➜ PDF":
                # Convert one image or multiple images → PDF
                image = Image.open(uploaded_file)
                pdf_buffer = BytesIO()
                rgb_image = image.convert("RGB")
                rgb_image.save(pdf_buffer, format="PDF", resolution=100.0)
                pdf_buffer.seek(0)
                st.download_button("⬇️ Download PDF", data=pdf_buffer, file_name="converted.pdf", mime="application/pdf")

        except Exception as e:
            st.error(f"❌ Conversion failed: {e}")
