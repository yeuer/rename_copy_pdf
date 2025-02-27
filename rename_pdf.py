import os
import fitz  # PyMuPDF
import re  # Untuk regex
import shutil  # Untuk copy file
import pandas as pd  # Untuk menyimpan log ke Excel
from datetime import datetime
import sys  # Untuk mendapatkan perintah eksekusi

# GANTI YANG INI (enaknya sih ke folder download)
folder_path = r"C:\selenium"  # Folder asal file PDF
output_folder = r"C:\selenium\hasil"  # Folder tujuan hasil copy
# GANTI SAMPAI INI SAJA

log_file = os.path.join(os.path.dirname(__file__), "log.xlsx")  # File log di folder yang sama

def extract_reference_and_name(pdf_path):
    """ Ekstrak nomor referensi dan nama pembeli dari file PDF """
    doc = fitz.open(pdf_path)
    reference_number = None
    buyer_name = None

    for page in doc:
        text = page.get_text("text")
        
        # Bersihkan teks dari spasi ekstra
        lines = [line.strip() for line in text.split("\n")]

        for line in lines:
            # Cari nomor referensi
            ref_match = re.search(r"referensi\s*:\s*(\d+)", line, re.IGNORECASE)
            if ref_match:
                reference_number = ref_match.group(1)

            # Cari nama pembeli (pastikan bukan "Pengusaha Kena Pajak")
            name_match = re.search(r"nama\s*:\s*(.+)", line, re.IGNORECASE)
            if name_match and "pengusaha kena pajak" not in line.lower():
                buyer_name = name_match.group(1).strip()

    return reference_number, buyer_name

def save_log_to_excel(command, logs):
    """ Simpan log ke file Excel """
    df = pd.DataFrame(logs, columns=["Perintah", "Tanggal & Waktu", "File Awal", "File Disimpan"])
    
    # Jika file Excel sudah ada, tambahkan data baru
    if os.path.exists(log_file):
        df_existing = pd.read_excel(log_file)
        df = pd.concat([df_existing, df], ignore_index=True)

    df.to_excel(log_file, index=False)
    print(f"📂 Log disimpan ke {log_file}")

def rename_and_copy_files(folder_path, output_folder):
    """ Rename file sesuai nomor referensi dan copy ke folder berdasarkan nama pembeli """
    logs = []
    command = " ".join(sys.argv)  # Perintah eksekusi yang dijalankan
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Waktu eksekusi

    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            reference_number, buyer_name = extract_reference_and_name(file_path)

            if reference_number and buyer_name:
                new_filename = f"{reference_number}.pdf"
                buyer_folder = os.path.join(output_folder, buyer_name)

                # Buat folder jika belum ada
                if not os.path.exists(buyer_folder):
                    os.makedirs(buyer_folder)

                new_file_path = os.path.join(buyer_folder, new_filename)

                # Cut file
                shutil.move(file_path, new_file_path)
                log_message = f"✅ Berhasil Copy: {filename} → {new_file_path}"
                print(log_message)

                # Tambahkan ke log
                logs.append([command, timestamp, filename, new_file_path])

    # Simpan log ke Excel
    if logs:
        save_log_to_excel(command, logs)

# Jalankan fungsi
rename_and_copy_files(folder_path, output_folder)
