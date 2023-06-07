import os
from openpyxl import Workbook

directory = "./png"
png_files = [file for file in os.listdir(directory) if file.endswith(".png")]

workbook = Workbook()
sheet = workbook.active

for index, file_name in enumerate(png_files, start=1):
    sheet.cell(row=index, column=1, value=file_name)

excel_file_name = "tagged.xlsx"
workbook.save(excel_file_name)