import os
import sys
import openpyxl
import pandas as pd
from tika import parser

folder_name = "pdf_files" ## Type your folder name where pdf are located
wanted_metric = ["Toplam Katma Değer Vergisi","Matrah Toplamı"] ## Type wanted metrics
file_path = "C:\\Workspace\\PDF-to-Text\\pdf_files\\" ## Type your file path of your pdf folder
output_excel_file_name = "Text-Extraction-Result.xlsx" ## Type your wanted output file name

def find_value(pdf_file, wanted_metric):
    pdf_file_path = file_path + pdf_file

    wanted_metric_list = []
    for i in wanted_metric:
        raw = parser.from_file(pdf_file_path)
        value = raw['content'].split(i)[1].strip().split("\n")[0]
        wanted_metric_list.append(value)

    return wanted_metric_list

def list_files_in_folder(folder_path):
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    return files

file_name = list_files_in_folder(folder_name)

pdf_name_list = []
value_list = []

num = 0
for file in file_name:

    pdf_name = file

    inner_value_list = []
    for metric in wanted_metric:
        value = find_value(file, wanted_metric)
        inner_value_list.append(value)

    pdf_name_list.append(pdf_name)
    value_list.append(inner_value_list[num])
    num += 1

result = pd.DataFrame({"File Name": pdf_name_list})

num = 0
for val in value_list:
    name = wanted_metric[num]
    num += 1
    result[name] = val

result.to_excel(output_excel_file_name, index=False)

print(result)