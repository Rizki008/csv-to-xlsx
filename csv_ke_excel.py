
import csv
import os
from openpyxl import Workbook

folder_input = "hasil_split_september"
folder_output = "hasil_excel"

os.makedirs(folder_output, exist_ok=True)

for nama_file in os.listdir(folder_input):
    if not nama_file.lower().endswith(".csv"):
        continue

    path_csv = os.path.join(folder_input, nama_file)
    nama_excel = os.path.splitext(nama_file)[0] + ".xlsx"
    path_excel = os.path.join(folder_output, nama_excel)

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")

    with open(path_csv, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)

        for row in reader:
            ws.append(row)

    wb.save(path_excel)
    print(f"Berhasil: {nama_file} -> {nama_excel}")

print("Semua file selesai dikonversi.")
