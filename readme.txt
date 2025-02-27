//kode untuk rename sesuai referensi dari file download dari coretax dan cut ke folder yang sama dengan nama pajak coretax, missal nama PT ABC
// misal nama file sebelumnya bsdjhsbdkkfd.pdf, akan membaca referensi dulu, kemudian rename menjadi no referensi tersebut
// buat folder sesuai nama penerima pajak (kalau belum ada)
// cut file ke folder hasil/nama_pt/no referensi
// buat log ke excel


## cara instal ##
// instal python(python saat kode ini dibuat 3.10.6)
winget install Python.Python

// untuk baca pdf
pip install pymupdf

// untuk copy paste
pip install pandas openpyxl


// ganti sesuai kebutuhan, file ada dimana dan mau cut kemana
# GANTI YANG INI
folder_path = r"C:\selenium"  # Folder asal file PDF
output_folder = r"C:\selenium\hasil"  # Folder tujuan hasil copy
# GANTI SAMPAI INI SAJA

## cara running ##
// buka cmd, cd ke tempat simpan file python
cd..
cd C
cd selenium

// run perintah untuk menjalankan file python
python rename_pdf.py

// cek hasil, pastikan ceklis ijo
// cek hasil yg sukses di log.xlsx

DONE