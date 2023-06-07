import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

script_directory = os.path.dirname(os.path.abspath(__file__))

folder_name = "png"

folder_path = os.path.join(script_directory, folder_name)

png_files = [file for file in os.listdir(folder_path) if file.endswith(".png")]

#for future use
time = [file[8:-9] if file.endswith("_Full.png") else file for file in png_files]

workbook = Workbook()
sheet = workbook.active

sheet.merge_cells("A1:J2")

merged_value = "Nama File Label Img Hari /Bulan / Tahun"
sheet["A1"].value = merged_value

sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")

header_values = [
    "Nomor",
    "Nama Foto",
    "Bacaan Plat Kiri",
    "Pelat Tertutup Kirim",
    "Pelat Tidak Jelas Kiri",
    "Bacaan Pelat Kanan",
    "Pelat Tertutup Kanan",
    "Pelat Tidak Jelas Kanan",
    "Xml ADA / TIDAK_ADA",
    "Jam PAGI / MALAM"
]

for col, value in enumerate(header_values, start=1):
    column_letter = get_column_letter(col)
    sheet[column_letter + "3"].value = value

column_width = 25
for col in range(1, len(header_values) + 1):
    column_letter = get_column_letter(col)
    sheet.column_dimensions[column_letter].width = column_width

border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
for row in sheet.iter_rows(min_row=1, max_row=3, max_col=len(header_values)):
    for cell in row:
        cell.border = border

for index, file_name in enumerate(png_files, start=1):
    number = index
    number += 0
    sheet.cell(row=index + 3, column=1, value=number)
    sheet.cell(row=index + 3, column=2, value=file_name)

# Add borders to cells in columns C to J based on column 1
for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row, min_col=1, max_col=len(header_values)):
    for cell in row:
        row_number = cell.row
        column_number = cell.column

        # Get the value from column 1 for the current row
        column_1_value = sheet.cell(row=row_number, column=1).value

        # Check if column 1 value is not empty
        if column_1_value:
            # Apply borders to the cells in columns C to J
            cell.border = border

excel_file_name = "tagged.xlsx"
workbook.save(excel_file_name)