import openpyxl
import os
import glob

SOURCE_DIR = "indirilen_ekler"
OUTPUT_FILE = "Guncel_Master_Veri.xlsx"

def process_and_merge_files():
    files = glob.glob(os.path.join(SOURCE_DIR, "*.xlsx"))
    if len(files) < 2:
        print("İşlem için en az 2 dosya gerekiyor.")
        return

    # 1. Dosyaları yükle
    file_a = files[0]
    file_b = files[1]
    
    wb_a = openpyxl.load_workbook(file_a)
    ws_a = wb_a.active

    wb_b = openpyxl.load_workbook(file_b)
    ws_b = wb_b.active

    # 5. satırdaki başlıkları bul (Row 5)
    headers = [cell.value for cell in ws_a[5]]
    
    try:
        col_bcsl = headers.index("BCSL0018") + 1
        col_aazbn = headers.index("AAZBN00") + 1
    except ValueError as e:
        print(f"Hata: Sütun başlıkları bulunamadı! {e}")
        return

    # Hangi dosya kimin? (Hangi dosyada hangi sütun daha doluysa o günceldir)
    # Ya da basitçe: B dosyasındaki AAZBN00 verilerini A dosyasına kopyalayalım.
    # (Eğer A dosyası BCSL0018 için güncelse)
    
    print(f"Sütunlar bulundu: BCSL0018 (Kolon {col_bcsl}), AAZBN00 (Kolon {col_aazbn})")

    # B'den A'ya AAZBN00 sütununu aktar (6. satırdan itibaren)
    for row in range(6, ws_a.max_row + 1):
        val_b = ws_b.cell(row=row, column=col_aazbn).value
        # Sadece boş olmayanları veya tümünü aktarabiliriz
        ws_a.cell(row=row, column=col_aazbn).value = val_b

    # A'dan B'ye BCSL0018 sütununu aktar (Ya da tam tersi, ikisini tek dosyada birleştir)
    # Burada mantık: wb_a artık hem kendi BCSL0018'ine hem de B'den gelen AAZBN00'e sahip.

    wb_a.save(OUTPUT_FILE)
    print(f"Başarıyla birleştirildi ve '{OUTPUT_FILE}' olarak kaydedildi.")

if __name__ == "__main__":
    process_and_merge_files()
