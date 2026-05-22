import pandas as pd
from outlook_bot import check_and_download_specific_mails
from veri_isleyici import process_and_merge_files
import schedule
import time
import os
import glob

def update_visualization_data(excel_path):
    """
    Guncel_Master_Veri.xlsx dosyasını okur ve main_vis.py'ın 
    kullandığı jet_brent_current.csv dosyasını günceller.
    """
    print("\nAdım 3: Görselleştirme verisi (CSV) güncelleniyor...")
    try:
        # Excel'i oku (main_vis.py mantığına göre: 5. satır başlık, 6. satırdan veri başlar)
        # main_vis.py'da skiprows=4 yapılmıştı, yani 5. satır header.
        df_up = pd.read_excel(excel_path, skiprows=4)
        
        # main_vis.py'daki sütun eşleştirme mantığı
        # Sütun isimleri farklı gelebileceği için main_vis.py'daki "Unnamed: 0", "BCSL0018", "AAZBN00" mantığını kullanıyoruz
        df_up = df_up.iloc[:, [0, 1, 2]] # İlk 3 sütunu al (Tarih, Brent, Jet)
        df_up.columns = ["tarih", "brent", "cif med"]
        df_up = df_up.dropna(subset=["tarih"])
        df_up["tarih"] = pd.to_datetime(df_up["tarih"])

        # Mevcut CSV dosyasını bul (jet_brent_current.csv veya en yeni olan)
        csv_path = "jet_brent_current.csv"
        if os.path.exists(csv_path):
            df_main = pd.read_csv(csv_path)
            df_main["tarih"] = pd.to_datetime(df_main["tarih"])
            
            # Güncelleme mantığı (Aynı tarihler varsa üzerine yaz, yoksa ekle)
            for _, row in df_up.iterrows():
                if pd.isna(row['brent']) and pd.isna(row['cif med']): continue
                
                match = df_main.index[df_main['tarih'] == row['tarih']].tolist()
                if match:
                    idx = match[0]
                    if pd.notna(row['brent']): df_main.at[idx, 'brent'] = row['brent']
                    if pd.notna(row['cif med']): df_main.at[idx, 'cif med'] = row['cif med']
                else:
                    df_main = pd.concat([df_main, pd.DataFrame([row])], ignore_index=True)
            
            df_main = df_main.sort_values("tarih")
            df_main.to_csv(csv_path, index=False)
            print(f"BAŞARILI: {csv_path} dosyası yeni verilerle güncellendi.")
        else:
            df_up.to_csv(csv_path, index=False)
            print(f"YENİ: {csv_path} dosyası oluşturuldu.")
            
    except Exception as e:
        print(f"Görselleştirme güncelleme hatası: {e}")

def run_bot_logic(sender_x, sender_y, folder_name):
    print(f"\n--- KONTROL BAŞLATILDI ({time.strftime('%H:%M:%S')}) ---")
    
    # 1. Adım: Mailleri kontrol et ve indir
    success = check_and_download_specific_mails(sender_x, sender_y, folder_name=folder_name)
    
    if success:
        # 2. Adım: Excel dosyalarını birleştir
        print("Adım 2: Veriler birleştiriliyor (BCSL from X, AAZBN from Y)...")
        process_and_merge_files()
        
        # 3. Adım: Görselleştirme CSV'sini otomatik besle
        update_visualization_data("Guncel_Master_Veri.xlsx")
        
        print("--- TÜM SÜREÇ BAŞARIYLA TAMAMLANDI ---\n")
        return True
    else:
        print(f"Eksik mail: Bugün her iki kişiden de mail gelmesi bekleniyor.")
        return False

def scheduled_job(sender_x, sender_y, folder_name):
    completed = run_bot_logic(sender_x, sender_y, folder_name)
    if not completed:
        def retry():
            nonlocal completed
            completed = run_bot_logic(sender_x, sender_y, folder_name)
            if completed:
                return schedule.CancelJob
        schedule.every(15).minutes.do(retry)

def main():
    print("--- JET FUEL ÖZEL BİRLEŞTİRME BOTU ---")
    
    folder_name = input("Outlook klasör adı (Varsayılan: jet fuel): ") or "jet fuel"
    
    sender_x = input("X Kişisi (BCSL güncelleyen) mail adresi: ")
    sender_y = input("Y Kişisi (AAZBN güncelleyen) mail adresi: ")
    
    run_time = input("Her gün çalışacağı saat (Örn: 09:30): ")
    
    schedule.every().day.at(run_time).do(scheduled_job, sender_x=sender_x, sender_y=sender_y, folder_name=folder_name)
    print(f"\nBot kuruldu! X:{sender_x} ve Y:{sender_y} mailleri beklenecek.")
    
    if input("Hemen şimdi kontrol edilsin mi? (e/h): ").lower() == 'e':
        run_bot_logic(sender_x, sender_y, folder_name)

    while True:
        schedule.run_pending()
        time.sleep(30)

if __name__ == "__main__":
    main()
