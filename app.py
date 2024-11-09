import streamlit as st
import re
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
import shutil
import datetime

# Path ke file Excel konfigurasi dan file bundel
AMM_REF_TO_TASK_CARD_PATH = "/content/AMM REF TO TASK CARD.xlsx"
REGISTRATION_TO_CONFIG_CODE_PATH = "/content/REGISTRATION TO CONFIG.xlsx"
BUNDLES = {
    "AWW": "/content/PDF BUNDEL (AWW).pdf",
    "CGP": "/content/PDF BUNDEL (CGP).pdf",
    "CHI": "/content/PDF BUNDEL (CHI).pdf",
    "GEF": "/content/PDF BUNDEL (GEF).pdf",
    "HAZ": "/content/PDF BUNDEL (HAZ).pdf",
    "ILF": "/content/PDF BUNDEL (ILF).pdf",
    "LOM": "/content/PDF BUNDEL (LOM).pdf",
    "OMR_SAOC": "/content/PDF BUNDEL (OMR_SAOC).pdf",
    "TCI": "/content/PDF BUNDEL (TCI).pdf",
    "GIA": "/content/PDF BUNDEL GIA.pdf",

    # Tambahkan konfigurasi bundel lainnya jika ada
}

# Fungsi untuk memuat data Task Card dari file Excel
def load_task_card_data():
    df = pd.read_excel(AMM_REF_TO_TASK_CARD_PATH)
    return {row['AMM REF']: row['Nomor Task Card'] for _, row in df.iterrows()}

# Fungsi untuk memuat data konfigurasi dari file Excel
def load_configuration_data():
    df = pd.read_excel(REGISTRATION_TO_CONFIG_CODE_PATH)
    return {row['Nomor Registrasi']: row['Kode Konfigurasi'] for _, row in df.iterrows()}

# Fungsi untuk menemukan nomor Task Card berdasarkan AMM REF di dokumen order
def find_task_card_number(order_doc, task_card_data):
    order_reader = PdfReader(order_doc)
    for page in order_reader.pages:
        text = page.extract_text()
        match = re.search(r"AMM REF\.:([\d-]+)", text)
        if match:
            amm_ref = match.group(1).strip()
            return task_card_data.get(amm_ref, None)
    return None

# Fungsi untuk menemukan nomor registrasi dalam dokumen order
def find_registration_number(order_doc):
    order_reader = PdfReader(order_doc)
    for page in order_reader.pages:
        text = page.extract_text()
        match = re.search(r"(PK-\w+)", text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None

# Fungsi untuk mendapatkan kode konfigurasi berdasarkan nomor registrasi
def find_configuration_code(registration_number, config_data):
    return config_data.get(registration_number, None)

# Fungsi untuk menambahkan watermark tanggal pada setiap halaman PDF
def add_watermark(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    today_date = datetime.datetime.today().strftime('%Y-%m-%d')
    watermark_pdf_path = "/content/temp_watermark.pdf"
    c = canvas.Canvas(watermark_pdf_path, pagesize=letter)
    c.setFont("Helvetica", 8)
    c.drawString(500, 10, f"Created on: {today_date}")  # Posisi di bawah kanan
    c.save()

    watermark_reader = PdfReader(watermark_pdf_path)
    watermark_page = watermark_reader.pages[0]

    for page in reader.pages:
        page.merge_page(watermark_page)
        writer.add_page(page)

    with open(output_pdf, "wb") as out_file:
        writer.write(out_file)

# Fungsi untuk memisahkan Task Card dari PDF bundel
def split_task_card(pdf_path, task_card_number, output_folder):
    pages = find_task_card_pages(pdf_path, task_card_number)
    if not pages:
        st.warning(f"Task card {task_card_number} tidak ditemukan di PDF bundel.")
        return None

    writer = PdfWriter()
    reader = PdfReader(pdf_path)
    for page_num in pages:
        writer.add_page(reader.pages[page_num])

    # Tentukan nama file output
    output_pdf = f"{output_folder}/{task_card_number}_extracted.pdf"
    with open(output_pdf, "wb") as out_file:
        writer.write(out_file)

    # Menambahkan watermark
    watermark_output_pdf = f"{output_folder}/{task_card_number}_watermarked.pdf"
    add_watermark(output_pdf, watermark_output_pdf)

    # Hapus file sementara tanpa watermark
    os.remove(output_pdf)

    return watermark_output_pdf if os.path.exists(watermark_output_pdf) else None

  # Fungsi untuk menemukan halaman-halaman Task Card di dalam PDF bundel berdasarkan nomor Task Card
def find_task_card_pages(pdf_path, task_card_number):
    pages_with_task_card = []
    reader = PdfReader(pdf_path)
    for page_num, page in enumerate(reader.pages):
        text = page.extract_text()
        if task_card_number in text:
            pages_with_task_card.append(page_num)
    return pages_with_task_card

# Fungsi untuk menggabungkan dokumen order dengan Task Card
def merge_order_with_task_card(order_pdf_path, task_card_pdf_path, output_path):
    writer = PdfWriter()
    order_reader = PdfReader(order_pdf_path)
    for page in order_reader.pages:
        writer.add_page(page)

    task_card_reader = PdfReader(task_card_pdf_path)
    for page in task_card_reader.pages:
        writer.add_page(page)

    with open(output_path, "wb") as out_file:
        writer.write(out_file)
    st.success(f"Penggabungan selesai. Hasil disimpan di {output_path}")

# Fungsi untuk membagi PDF berdasarkan nomor order
def split_pdf_by_order(uploaded_file, output_folder):
    reader = PdfReader(uploaded_file)
    total_pages = len(reader.pages)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    writer = PdfWriter()
    order_number = None
    output_path = None

    for i in range(total_pages):
        page = reader.pages[i]
        page_text = page.extract_text()

        match = re.search(r"ORDER\s*[:]*\s*(\d+)", page_text)
        if match:
            new_order_number = match.group(1)

            if order_number is not None and order_number != new_order_number:
                if len(writer.pages) > 0:
                    output_path = f"{output_folder}/order_{order_number}.pdf"
                    with open(output_path, "wb") as outfile:
                        writer.write(outfile)

                writer = PdfWriter()

            order_number = new_order_number
            output_path = f"{output_folder}/order_{order_number}.pdf"

        writer.add_page(page)

    if len(writer.pages) > 0:
        output_path = f"{output_folder}/order_{order_number}.pdf"
        with open(output_path, "wb") as outfile:
            writer.write(outfile)

# Fungsi untuk menggabungkan semua file PDF hasil penggabungan ke satu file PDF
def merge_all_pdfs(pdf_files, final_output_path):
    writer = PdfWriter()
    for pdf_file in pdf_files:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            writer.add_page(page)

    with open(final_output_path, "wb") as out_file:
        writer.write(out_file)

# Streamlit UI
def main():
    st.title("JOB CARD GENERATOR")
    st.markdown("### Integrating Order and Maintenance Manual Extract")

    dokumen_order = st.file_uploader("Upload Order Document", type="pdf")

    if dokumen_order:
        task_card_data = load_task_card_data()
        config_data = load_configuration_data()

        output_folder = "/content/split_orders"
        split_pdf_by_order(dokumen_order, output_folder)

        output_files = []

        for order_file in os.listdir(output_folder):
            order_file_path = os.path.join(output_folder, order_file)
            task_card_number = find_task_card_number(order_file_path, task_card_data)
            if not task_card_number:
                continue

            registration_number = find_registration_number(order_file_path)
            if not registration_number:
                continue

            config_code = find_configuration_code(registration_number, config_data)
            if not config_code:
                continue

            task_card_pdf = None
            if config_code in BUNDLES:
                task_card_pdf = split_task_card(BUNDLES[config_code], task_card_number, output_folder)

            if task_card_pdf:
                output_pdf = f"merged_{order_file.replace('.pdf', '')}_{task_card_number}.pdf"
                merge_order_with_task_card(order_file_path, task_card_pdf, output_pdf)
                output_files.append(output_pdf)

        final_output_pdf = "final_merged_output.pdf"
        merge_all_pdfs(output_files, final_output_pdf)

        with open(final_output_pdf, "rb") as f:
            st.download_button("Download Final PDF", f, file_name=final_output_pdf)

if __name__ == "__main__":
    main()

        
