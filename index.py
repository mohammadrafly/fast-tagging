import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Get the current script's directory
script_directory = os.path.dirname(os.path.abspath(__file__))

# Specify the folder name containing the PNG files
folder_name = "png"

# Create the full path to the folder
folder_path = os.path.join(script_directory, folder_name)

png_files = [file for file in os.listdir(folder_path) if file.endswith(".png")]

# Create an Excel Workbook
workbook = Workbook()
sheet = workbook.active

# Merge Cells A1:J2
sheet.merge_cells("A1:J2")

# Set Merged Cell Value
merged_value = "Nama File Label Img Hari /Bulan / Tahun"
sheet["A1"].value = merged_value

# Set Center Alignment for Merged Cells A1:J2
sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")

# Set Column Headers
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

# Set fixed width for columns A3 to J3
column_width = 25  # Adjust this value as per your preference
for col in range(1, len(header_values) + 1):
    column_letter = get_column_letter(col)
    sheet.column_dimensions[column_letter].width = column_width

# Add borders to cells
border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
for row in sheet.iter_rows(min_row=1, max_row=3, max_col=len(header_values)):
    for cell in row:
        cell.border = border

# Write PNG File Names and "taken" to Excel
for index, file_name in enumerate(png_files, start=1):
    sheet.cell(row=index + 3, column=2, value=file_name)

# Add borders to cells with data
for row in sheet.iter_rows(min_row=4, max_row=4, max_col=len(header_values)):
    for cell in row:
        if cell.value:
            cell.border = border

# Save the Excel File
excel_file_name = "tagged.xlsx"  # Change this to your desired file name
workbook.save(excel_file_name)
