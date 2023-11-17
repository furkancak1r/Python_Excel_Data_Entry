import win32com.client
import json
import questionary
while True:
    try:        
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.ActiveWorkbook
        workbook_name = workbook.Name
        print(workbook.Name)

        # Excel uygulamasının görünür olmasını sağla
        excel.Visible = True

        # rot.json'dan rota adlarını ve kodlarını sıralı bir şekilde oku
        with open("rot.json", "r", encoding="utf-8") as rot_file:
            rot_data = sorted(list(json.load(rot_file).items()))

        # mamul.json'dan verileri oku
        with open("mamul.json", "r", encoding="utf-8") as mamul_file:
            mamul_data = json.load(mamul_file)

        # Excel sayfa seçenekleri
        sheet_options = [sheet.Name for sheet in workbook.Worksheets]
        sheet_options.append("Çıkış") 

        # Kullanıcıdan sayfa seçmesini iste
        sheet_question = {
            'type': 'select',
            'name': 'selected_sheet',
            'message': 'Düzenlemek istediğiniz sayfayı seçin:',
            'choices': sheet_options,
        }
        selected_sheet = questionary.prompt(sheet_question)['selected_sheet']

        if selected_sheet == "Çıkış":
            print("Programdan çıkılıyor.")
            break
        else:
            # Rot.json'daki rota adlarını ve kodlarını göster ve kullanıcıdan seçmeyi iste
            selected_rot_question = {
                'type': 'select',
                'name': 'selected_rot',
                'message': 'Yeni sayfa adını seçin:',
                'choices': [f"{rot_code} - {rot_name}" for rot_code, rot_name in rot_data],
            }
            # Seçilen rota kodunun sadece sayısal kısmını al
            selected_rot = questionary.prompt(selected_rot_question)['selected_rot']
            selected_rot = int(selected_rot.split()[0][3:]) - 1

            # Seçilen rota adını yeni sayfa adı olarak kullan
            workbook.Worksheets(selected_sheet).Name = rot_data[selected_rot][1]

            cell_start = input("Başlangıç hücresini girin: ")

            
            sheet = workbook.Worksheets(rot_data[selected_rot][1])

            max_row = sheet.UsedRange.Rows.Count
            # Determine the column of the cell_start (assuming it is in the format "A1", "B2", etc.)
            start_column = ord(cell_start[0].upper()) - ord('A') + 1
            # Dynamically create cell_end using the max_row and start_column
            cell_end = f"{chr(start_column + ord('A') - 1)}{max_row+1}"
            cell_left_count = int(input("Solda kaç hücre var: "))

            with open("mamul.json", "r", encoding="utf-8") as f:
                json_data = json.load(f)

            variant_index = workbook_name.lower().find("varyant")
            workbook_prefix = workbook_name[:variant_index].strip()

            # Eşleşen seçeneği bul ve sonucu yazdır
            for level, categories in json_data.items():
                for category, variants in categories.items():
                    for variant, variant_code in variants.items():
                        if variant_code.lower() == workbook_prefix.lower():
                            print(f"{workbook_prefix} içindeki eşleşen numaralar: {level} ve {variant}")

                            result = str(level) + str(variant)

                            # Kullanıcının girdiği sayfayı seç
                            sheet = workbook.Worksheets(rot_data[selected_rot][1])

                            # Başlangıç ve bitiş hücrelerini bir aralık olarak tanımla
                            cell_range = sheet.Range(cell_start, cell_end)                

                            # Aralıktaki tüm hücreleri döngü ile gez
                            # Bir sayacı tanımla ve 1'den başlat
                            counter = 1
                            for cell in cell_range:
                                # Hücrenin değerini sonuca eşitle
                                # Sayacı 4 basamaklı bir string olarak formatla ve sonuca ekle
                                # Seçilen rota kodunu da sonuca ekle
                                cell.Value = result + f"{counter:04d}-{rot_data[selected_rot][0]}"
                                adjacent_cell = sheet.Cells(cell.Row, cell.Column + 1)
                                adjacent_cell.Value = f"{workbook_prefix} {rot_data[selected_rot][1]}"
                                # Hücrenin solundaki hücre sayısı kadar ters döngü yap
                                for i in reversed(range(cell_left_count)):
                                    # Hücrenin solundaki hücreyi bul
                                    left_cell = sheet.Cells(cell.Row, cell.Column - (i + 1))
                                    # Hücrenin değerine solundaki hücrenin değerini de ekle
                                    adjacent_cell.Value += " " + str(left_cell.Value)
                                # Sayacı 1 arttır
                                counter += 1
                            used_range = sheet.UsedRange
                            used_range.EntireRow.AutoFit()
                            used_range.EntireColumn.AutoFit()  
            print("İşlem tamamlandı. Program devam ediyor...")
    except Exception as e:
            print(f'Hata oluştu: {e}')

   