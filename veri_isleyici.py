import pandas as pd
import os
import glob

# --- AYARLAR ---
SOURCE_DIR = "indirilen_ekler"
MASTER_FILE = "ana_veriler.xlsx"

def process_excel_files():
    # İndirilen ekler klasöründeki excel dosyalarını bul
    excel_files = glob.glob(os.path.join(SOURCE_DIR, "*.xlsx")) + glob.glob(os.path.join(SOURCE_DIR, "*.xls"))
    
    if not excel_files:
        print("İşlenecek yeni dosya bulunamadı.")
        return

    all_data = []
    
    # Mevcut ana dosya varsa oku (tekrarı önlemek için veya üzerine eklemek için)
    if os.path.exists(MASTER_FILE):
        master_df = pd.read_excel(MASTER_FILE)
        all_data.append(master_df)
        print(f"Mevcut {MASTER_FILE} dosyası yüklendi.")
    else:
        master_df = pd.DataFrame()
        print("Yeni ana dosya oluşturulacak.")

    for file in excel_files:
        try:
            print(f"İşleniyor: {file}")
            # Excel'i oku
            df = pd.read_excel(file)
            
            # Buraya özel sütun seçme mantığı eklenebilir
            # df = df[["Tarih", "Müşteri", "Tutar"]] gibi
            
            all_data.append(df)
            
            # İşlenen dosyayı arşive taşıma veya silme (isteğe bağlı)
            # os.rename(file, os.path.join("arsiv", os.path.basename(file)))
            
        except Exception as e:
            print(f"Dosya okunurken hata oluştu ({file}): {e}")

    if all_data:
        # Tüm verileri birleştir
        final_df = pd.concat(all_data, ignore_index=True)
        
        # Eğer tamamen aynı satırlar varsa temizle
        final_df = final_df.drop_duplicates()
        
        # Ana dosyaya kaydet
        final_df.to_excel(MASTER_FILE, index=False)
        print(f"İşlem tamamlandı! Toplam satır sayısı: {len(final_df)}")
        
        # İşlenen dosyaları temizle (isteğe bağlı - her seferinde aynı veriyi çekmemek için)
        # cleanup_source_dir(excel_files)

def cleanup_source_dir(files):
    for f in files:
        os.remove(f)
    print("İşlenen kaynak dosyalar temizlendi.")

if __name__ == "__main__":
    process_excel_files()
