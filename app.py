import streamlit as st
from io import BytesIO
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pdf2image import convert_from_bytes
from PIL import Image
from PyPDF2 import PdfReader

st.set_page_config(page_title="File Converter", page_icon="📄", layout="centered")

st.title("📄 Universal File Converter")
st.markdown("Convert between Word, PDF, and Image formats easily!")

uploaded_file = st.file_uploader("📤 Upload your file", type=["pdf", "docx", "jpg", "jpeg", "png"])

conversion_type = st.selectbox(
    "🔄 Select conversion type",
    ["Select", "Word ➜ PDF", "PDF ➜ Word", "PDF ➜ Image", "Image ➜ PDF"]
)

if uploaded_file and conversion_type != "Select":
    if st.button("🚀 Convert"):
        try:
            if conversion_type == "Word ➜ PDF":
                # Convert Word to PDF
                doc = Document(uploaded_file)
                pdf_buffer = BytesIO()
                c = canvas.Canvas(pdf_buffer, pagesize=letter)
                text = c.beginText(50, 750)
                for para in doc.paragraphs:
                    text.textLine(para.text)
                c.drawText(text)
                c.save()
                st.download_button("⬇️ Download PDF", data=pdf_buffer.getvalue(),
                                   file_name="converted.pdf", mime="application/pdf")

            elif conversion_type == "PDF ➜ Word":
                # Convert PDF to Word
                pdf = PdfReader(uploaded_file)
                doc = Document()
                for page in pdf.pages:
                    doc.add_paragraph(page.extract_text() or "")
                buffer = BytesIO()
                doc.save(buffer)
                st.download_button("⬇️ Download Word", data=buffer.getvalue(),
                                   file_name="converted.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            elif conversion_type == "PDF ➜ Image":
                # Convert PDF to Images
                images = convert_from_bytes(uploaded_file.read())
                for i, img in enumerate(images):
                    img_buffer = BytesIO()
                    img.save(img_buffer, format="PNG")
                    st.image(img, caption=f"Page {i+1}")
                    st.download_button(f"⬇️ Download Page {i+1}",
                                       data=img_buffer.getvalue(),
                                       file_name=f"page_{i+1}.png",
                                       mime="image/png")

            elif conversion_type == "Image ➜ PDF":
                # Convert Image to PDF
                image = Image.open(uploaded_file)
                pdf_buffer = BytesIO()
                image.convert("RGB").save(pdf_buffer, format="PDF")
                st.download_button("⬇️ Download PDF", data=pdf_buffer.getvalue(),
                                   file_name="converted.pdf", mime="application/pdf")

        except Exception as e:
            st.error(f"❌ Conversion failed: {e}")
