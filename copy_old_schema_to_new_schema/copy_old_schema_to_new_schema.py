import os
import json
from openpyxl import load_workbook

new_excel_directory = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Python_Excel_Data_Entry\\Şablonlar_rev01\\yeni_exceller"
old_excel_directory = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Python_Excel_Data_Entry\\Şablonlar_rev01\\old_exceller"
json_file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Python_Excel_Data_Entry\\copy_old_schema_to_new_schema\\excel_names_2.json"

with open(json_file_path, "r", encoding="utf-8") as json_file:
    excel_names_json = json.load(json_file)

new_excel_files = {file for file in os.listdir(
    new_excel_directory) if file.endswith('.xlsx')}
modified_excel_names = ["PROSO KALEM ANA VERİSİ ÖRNEK ŞABLONU_" +
                        name + ".xlsx" for name in excel_names_json]

matching_files = []
found_cells = []

for old_excel_file in os.listdir(old_excel_directory):
    if old_excel_file and not old_excel_file.startswith("~$"):
        old_excel_file_path = os.path.join(old_excel_directory, old_excel_file)

        if os.path.exists(old_excel_file_path):
            old_workbook = load_workbook(old_excel_file_path, read_only=True)

            old_sheet_names = {sheet.lower()
                               for sheet in old_workbook.sheetnames}

            for sheet_name in excel_names_json:
                try:
                    sheet_name_lower = sheet_name.lower()

                    if sheet_name_lower in old_sheet_names:
                        sheet_workbook = old_workbook[sheet_name]
                        search_values = ["MODUL", "MODÜL"]
                        for row_index, row in enumerate(sheet_workbook.iter_rows(values_only=True), start=1):
                            kodlar_values = []  # Initialize kodlar_values for each row
                            for col_index, cell_value in enumerate(row, start=1):
                                if cell_value:
                                    cell_value_lower = str(cell_value).lower()
                                    if "kod" in cell_value_lower or "kodlar" in cell_value_lower:
                                        found_cells.append({
                                            "sheet_name": sheet_name,
                                            "cell_value": cell_value,
                                            "row_index": row_index,
                                            "col_index": col_index
                                        })
                                        if sheet_name =="ELEKTRİK MONTAJ": #burası kaldırılacak gerçekte
                                            for row_index_2 in range(row_index + 1, sheet_workbook.max_row + 1):
                                                print("sheet_name,row_index_2:", sheet_name,row_index_2)
                                                kodlar_values.append(sheet_workbook.cell(row=row_index_2, column=col_index).value)
                                            if kodlar_values:
                                                found_cells[-1]["kodlar_values"] = kodlar_values

                except Exception as e:
                    print(f"Error opening workbook: {e}")

            print("found_cells:", found_cells)
