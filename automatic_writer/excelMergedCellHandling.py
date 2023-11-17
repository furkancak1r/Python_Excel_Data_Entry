# win32com.client modülünü içe aktar
import win32com.client

# Bir excel uygulaması nesnesi oluştur
excel = win32com.client.Dispatch("Excel.Application")

# Kaynak klasördeki tüm xlsx dosyalarını listele
import os
source_folder = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\python_sap_varyant\\1" # Kaynak klasör yolu
target_folder = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\python_sap_varyant\\2" # Hedef klasör yolu
xlsx_files = [f for f in os.listdir(source_folder) if f.endswith(".xlsx")]

# Her bir xlsx dosyası için
for xlsx_file in xlsx_files:
    # Dosyayı aç
    workbook = excel.Workbooks.Open(os.path.join(source_folder, xlsx_file))
    
    # Vba kodunun pythona çevrilmiş halini çalıştır
    for ws in workbook.Worksheets: # Her bir çalışma sayfası için
        for cell in ws.UsedRange: # Kullanılan aralıktaki her bir hücre için
            if cell.MergeCells: # Eğer hücre birleştirilmişse
                mergedCells = cell.MergeArea # Birleştirilmiş hücreleri al
                cell.MergeCells = False # Birleştirmeyi kaldır
                mergedCells.Value = cell.Value # Birleştirilmiş hücrelere aynı değeri ata
    
    # Dosyayı hedef klasöre xlsx formatında kaydet
    workbook.SaveAs(os.path.join(target_folder, xlsx_file), FileFormat=51)
    
    # Dosyayı kapat
    workbook.Close()
    
    # Biten excelin ismini konsola yaz
    print(f"{xlsx_file} dosyası {target_folder} klasörüne kaydedildi.")
    
# Excel uygulaması nesnesini kapat ve sil
excel.Quit()
