import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
import re
import io

st.title("Sistem Split Dokumen Berdasarkan Nomor Order")

def split_pdf_by_order(uploaded_file):
    if not uploaded_file:
        st.warning("Silakan unggah file bundle PDF.")
        return None

    pdf_reader = PdfReader(uploaded_file)
    num_pages = len(pdf_reader.pages)
    order_pattern = re.compile(r"ORDER : (\d+)")
    documents = []
    current_writer = PdfWriter()
    current_order = None

    for page_num in range(num_pages):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()

        match = order_pattern.search(text)
        if match:
            if current_order:
                documents.append((current_order, current_writer))
                current_writer = PdfWriter()
            current_order = match.group(1)
        
        current_writer.add_page(page)

    if current_order:
        documents.append((current_order, current_writer))

    return documents

def download_split_files(documents):
    for order_id, writer in documents:
        output = io.BytesIO()
        writer.write(output)
        output.seek(0)
        st.download_button(
            label=f"Unduh Order {order_id}",
            data=output,
            file_name=f"order_{order_id}.pdf",
            mime="application/pdf"
        )

uploaded_file = st.file_uploader("Unggah file PDF bundle", type="pdf")

if uploaded_file:
    st.write("Memproses dokumen...")
    split_documents = split_pdf_by_order(uploaded_file)

    if split_documents:
        st.success("Dokumen berhasil dipisah!")
        download_split_files(split_documents)
    else:
        st.error("Tidak ditemukan dokumen yang sesuai.")
