from outlook_bot import check_and_download_specific_mails
from veri_isleyici import process_and_merge_files
import schedule
import time
import os

def run_bot_logic(sender_x, sender_y, folder_name):
    print(f"\n--- KONTROL BAŞLATILDI ({time.strftime('%H:%M:%S')}) ---")
    
    # Mailleri kontrol et ve X/Y olarak indir
    success = check_and_download_specific_mails(sender_x, sender_y, folder_name=folder_name)
    
    if success:
        print("Adım 2: Veriler birleştiriliyor (BCSL from X, AAZBN from Y)...")
        process_and_merge_files()
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
