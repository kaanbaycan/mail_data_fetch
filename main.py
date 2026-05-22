from outlook_bot import check_and_download_specific_mails
from veri_isleyici import process_excel_files
import schedule
import time
import os

def run_bot_logic(senders, keywords, folder_name):
    print(f"\n--- KONTROL BAŞLATILDI ({time.strftime('%H:%M:%S')}) ---")
    
    # 1. Adım: Belirlenen klasörde mailleri kontrol et
    success = check_and_download_specific_mails(senders, keywords, folder_name=folder_name)
    
    if success:
        # 2. Adım: İndirilen dosyaları işle
        print("Adım 2: Veriler ana dosyaya işleniyor...")
        process_excel_files()
        print("--- TÜM SÜREÇ BAŞARIYLA TAMAMLANDI ---\n")
        return True
    else:
        print(f"Eksik mail olduğu için işlem yapılamadı. {folder_name} klasörü kontrol edilmeye devam edilecek.")
        return False

def scheduled_job(senders, keywords, folder_name):
    completed = run_bot_logic(senders, keywords, folder_name)
    
    if not completed:
        def retry():
            nonlocal completed
            completed = run_bot_logic(senders, keywords, folder_name)
            if completed:
                return schedule.CancelJob

        schedule.every(15).minutes.do(retry)

def main():
    print("--- ÖZEL MAİL VERİ ÇEKME BOTU ---")
    
    # Klasör adını al
    folder_name = input("Outlook'taki klasör adını girin (Varsayılan için Enter'a basın - jet fuel): ")
    if not folder_name:
        folder_name = "jet fuel"

    # Gönderenleri al
    sender_input = input("Mail beklediğiniz kişilerin e-posta adreslerini girin (Virgülle ayırın): ")
    senders = [s.strip() for s in sender_input.split(",") if s.strip()]
    
    # Konu filtresi al
    subject_input = input("Maillerin konusunda geçmesi gereken ortak kelimeleri girin: ")
    keywords = [k.strip() for k in subject_input.split(",") if k.strip()]

    run_time = input("Botun her gün ilk kontrolü yapacağı saati girin (Örn: 09:30): ")
    
    # Zamanlayıcıyı kur
    schedule.every().day.at(run_time).do(scheduled_job, senders=senders, keywords=keywords, folder_name=folder_name)
    print(f"\nBot kuruldu! '{folder_name}' klasörü taranacak.")
    
    # Manuel çalıştırma
    ilk_calisma = input("Hemen şimdi kontrol edilsin mi? (e/h): ")
    if ilk_calisma.lower() == 'e':
        run_bot_logic(senders, keywords, folder_name)

    while True:
        schedule.run_pending()
        time.sleep(30)

if __name__ == "__main__":
    main()
