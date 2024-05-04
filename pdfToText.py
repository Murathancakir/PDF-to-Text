import os
import sys
import openpyxl
import pandas as pd
from tika import parser

folder_name = "pdf_files"
wanted_string = "Toplam Katma Değer Vergisi"
file_path = "C:\\Workspace\\Pdf To Text\\pdf_files\\"
output_excel_file_name = "Text-Extraction-Result.xlsx"

def find_value(pdf_file):
    pdf_file_path = file_path + pdf_file

    raw = parser.from_file(pdf_file_path)
    value = raw['content'].split(wanted_string)[1].strip().split("\n")[0]
    return value

def list_files_in_folder(folder_path):
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    return files

file_name = list_files_in_folder(folder_name)

pdf_name_list = []
value_list = []

for file in file_name:
    pdf_name = file
    value = find_value(file)
    pdf_name_list.append(pdf_name)
    value_list.append(value)

result = pd.DataFrame({"File Name": pdf_name_list, "Value": value_list})
result.to_excel(output_excel_file_name)

print(result)