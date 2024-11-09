import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

# Path ke file Excel yang berisi data referensi AMM dan nomor Task Card
AMM_REF_TO_TASK_CARD_PATH = '/path/to/AMM REF TO TASK CARD.xlsx'

# Membaca data dari file Excel
def load_task_card_data():
    # Memuat data dari Excel dan mengonversinya menjadi dictionary
    df = pd.read_excel(AMM_REF_TO_TASK_CARD_PATH)
    return {row['AMM REF']: row['Nomor Task Card'] for _, row in df.iterrows()}

# Fungsi untuk memisahkan PDF berdasarkan data referensi AMM
def split_pdf_by_order(pdf_files, output_folder, task_card_data):
    # Memastikan folder output ada
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Proses untuk setiap file PDF
    for pdf_file in pdf_files:
        reader = PdfReader(pdf_file)
        writer = PdfWriter()
        total_pages = len(reader.pages)
        
        # Iterasi setiap halaman dalam PDF
        for page_num in range(total_pages):
            page = reader.pages[page_num]
            writer.add_page(page)

            # Mengekstrak teks atau informasi tertentu dari halaman, jika diperlukan
            # Misalnya, kita mencari nomor AMM REF di dalam teks halaman
            text = page.extract_text()
            for amm_ref, task_card in task_card_data.items():
                if amm_ref in text:  # Cek jika referensi AMM ada dalam teks halaman
                    # Menyimpan output PDF yang dipisahkan berdasarkan AMM REF
                    output_pdf_path = os.path.join(output_folder, f'{task_card}_page_{page_num + 1}.pdf')
                    with open(output_pdf_path, 'wb') as out_file:
                        writer.write(out_file)
                    print(f'Page {page_num + 1} untuk {task_card} disimpan.')
                    writer = PdfWriter()  # Reset writer untuk halaman berikutnya

    return output_folder

# Fungsi utama untuk menjalankan program
def main():
    # Folder output yang digunakan untuk menyimpan file PDF hasil pemisahan
    output_folder = "./split_orders"

    # Mendapatkan data task card dari Excel
    task_card_data = load_task_card_data()

    # Tentukan file PDF yang ingin diproses
    pdf_files = ['order1.pdf', 'order2.pdf']  # Ganti dengan file PDF yang sesuai

    # Memanggil fungsi untuk memisahkan PDF berdasarkan referensi AMM
    split_pdf_by_order(pdf_files, output_folder, task_card_data)

    print("Pemrosesan selesai.")

if __name__ == "__main__":
    main()
