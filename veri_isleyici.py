import openpyxl
import os

SOURCE_DIR = "indirilen_ekler"
OUTPUT_FILE = "Guncel_Master_Veri.xlsx"

def process_and_merge_files():
    path_x = os.path.join(SOURCE_DIR, "X_file.xlsx")
    path_y = os.path.join(SOURCE_DIR, "Y_file.xlsx")

    if not os.path.exists(path_x) or not os.path.exists(path_y):
        print("HATA: X veya Y dosyası bulunamadı. Lütfen maillerin indiğinden emin olun.")
        return

    print("Dosyalar birleştiriliyor...")

    # Dosyaları yükle
    wb_x = openpyxl.load_workbook(path_x)
    ws_x = wb_x.active

    wb_y = openpyxl.load_workbook(path_y)
    ws_y = wb_y.active

    # 5. satırdaki başlıkları bul
    headers_x = [cell.value for cell in ws_x[5]]
    headers_y = [cell.value for cell in ws_y[5]]

    try:
        col_bcsl = headers_x.index("BCSL0018") + 1
        col_aazbn = headers_y.index("AAZBN00") + 1
        print(f"Sütunlar bulundu: BCSL0018 (Kolon {col_bcsl}), AAZBN00 (Kolon {col_aazbn})")
    except ValueError as e:
        print(f"HATA: Sütun başlıkları 5. satırda bulunamadı! {e}")
        return

    # MANTIK: X dosyasını ana taslak olarak kullanacağız.
    # X dosyasındaki AAZBN00 sütununu, Y dosyasındakiyle güncelleyeceğiz.
    # (BCSL zaten X'te güncel olduğu için dokunmuyoruz)
    
    for row in range(6, ws_x.max_row + 1):
        # Y dosyasındaki güncel AAZBN değerini al
        val_y_aazbn = ws_y.cell(row=row, column=col_aazbn).value
        # X dosyasına (Ana dosya) yapıştır
        ws_x.cell(row=row, column=col_aazbn).value = val_y_aazbn

    # Sonucu kaydet
    wb_x.save(OUTPUT_FILE)
    print(f"\nİŞLEM BAŞARILI!")
    print(f"X'ten BCSL0018, Y'den AAZBN00 sütunları alındı.")
    print(f"Yeni dosya: {os.path.abspath(OUTPUT_FILE)}")

if __name__ == "__main__":
    process_and_merge_files()
