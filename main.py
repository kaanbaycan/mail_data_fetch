from outlook_bot import check_and_download_specific_mails
from veri_isleyici import process_and_merge_files
import schedule
import time
import os

def run_bot_logic(folder_name):
    print(f"\n--- KONTROL BAŞLATILDI ({time.strftime('%H:%M:%S')}) ---")
    
    # Bugünün son 2 mailini bekle ve indir
    success = check_and_download_specific_mails(folder_name=folder_name)
    
    if success:
        print("Adım 2: Veriler birleştiriliyor (Merge işlemi)...")
        process_and_merge_files()
        print("--- TÜM SÜREÇ BAŞARIYLA TAMAMLANDI ---\n")
        return True
    else:
        print(f"Eksik mail (Bugünlük 2 mail henüz gelmemiş). 15 dk sonra tekrar denenecek.")
        return False

def scheduled_job(folder_name):
    completed = run_bot_logic(folder_name)
    if not completed:
        def retry():
            nonlocal completed
            completed = run_bot_logic(folder_name)
            if completed:
                return schedule.CancelJob
        schedule.every(15).minutes.do(retry)

def main():
    print("--- JET FUEL EXCEL ÇEKME BOTU ---")
    
    folder_name = input("Outlook klasör adı (Varsayılan: jet fuel): ")
    if not folder_name:
        folder_name = "jet fuel"

    run_time = input("Her gün çalışacağı saat (Örn: 09:30): ")
    
    schedule.every().day.at(run_time).do(scheduled_job, folder_name=folder_name)
    print(f"\nBot kuruldu! '{folder_name}' klasöründe her gün saat {run_time}'da 2 mail aranacak.")
    
    if input("Hemen şimdi kontrol edilsin mi? (e/h): ").lower() == 'e':
        run_bot_logic(folder_name)

    while True:
        schedule.run_pending()
        time.sleep(30)

if __name__ == "__main__":
    main()
