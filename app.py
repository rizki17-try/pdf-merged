import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
import re
import io
import pandas as pd

st.title("Sistem Split dan Merge Dokumen Berdasarkan Nomor Order dan Task Card")

def split_pdf_by_order(uploaded_file):
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

def split_task_card_by_reference(uploaded_file):
    pdf_reader = PdfReader(uploaded_file)
    num_pages = len(pdf_reader.pages)
    task_card_pattern = re.compile(r"BOEING CARD NO\. (\S+)")
    page_pattern = re.compile(r"Page \d+ of (\d+)")
    documents = []
    current_writer = PdfWriter()
    current_task_card = None
    max_pages = 0

    for page_num in range(num_pages):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()

        task_card_match = task_card_pattern.search(text)
        page_match = page_pattern.search(text)
        
        if task_card_match:
            if current_task_card:
                documents.append((current_task_card, current_writer))
                current_writer = PdfWriter()
            current_task_card = task_card_match.group(1)
            max_pages = int(page_match.group(1)) if page_match else 0

        current_writer.add_page(page)

        if page_match and page_num + 1 == num_pages:
            documents.append((current_task_card, current_writer))

    if current_task_card:
        documents.append((current_task_card, current_writer))

    return documents

def load_reference_mapping(reference_file):
    df = pd.read_excel(reference_file)
    reference_mapping = dict(zip(df["AMM REF"], df["Nomor Task Card"]))
    return reference_mapping

def merge_documents(order_documents, task_documents, reference_mapping):
    merged_documents = []
    for order_id, order_writer in order_documents:
        amm_ref = re.search(r"\d{2}-\d{2}-\d{2}-\d{3}-\d{3}", order_id)
        if amm_ref and amm_ref.group(0) in reference_mapping:
            task_card_no = reference_mapping[amm_ref.group(0)]
            for task_id, task_writer in task_documents:
                if task_id == task_card_no:
                    merged_writer = PdfWriter()
                    for page in order_writer.pages:
                        merged_writer.add_page(page)
                    for page in task_writer.pages:
                        merged_writer.add_page(page)
                    merged_documents.append((f"{order_id}_{task_id}", merged_writer))
    return merged_documents

def download_merged_files(documents):
    for doc_id, writer in documents:
        output = io.BytesIO()
        writer.write(output)
        output.seek(0)
        st.download_button(
            label=f"Unduh Dokumen {doc_id}",
            data=output,
            file_name=f"{doc_id}.pdf",
            mime="application/pdf"
        )

# Upload file PDF dan Excel
order_file = st.file_uploader("Unggah file PDF bundle dokumen order", type="pdf")
task_card_file = st.file_uploader("Unggah file PDF bundle dokumen task card", type="pdf")
reference_file = st.file_uploader("Unggah file referensi (AMM REF TO TASK CARD.xlsx)", type="xlsx")

# Proses dokumen setelah upload
if order_file and task_card_file and reference_file:
    st.write("Memproses dokumen...")

    # Split dokumen berdasarkan nomor order dan task card
    order_documents = split_pdf_by_order(order_file)
    task_documents = split_task_card_by_reference(task_card_file)

    # Load referensi dari file Excel
    reference_mapping = load_reference_mapping(reference_file)

    # Merge dokumen yang sudah di-split berdasarkan referensi
    merged_documents = merge_documents(order_documents, task_documents, reference_mapping)

    if merged_documents:
        st.success("Dokumen berhasil dipisah dan digabung!")
        download_merged_files(merged_documents)
    else:
        st.error("Tidak ditemukan dokumen yang sesuai.")
