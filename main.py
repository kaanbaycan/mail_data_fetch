from outlook_bot import check_and_download_specific_mails
from veri_isleyici import process_excel_files
import schedule
import time
import os

def run_bot_logic(senders, keywords):
    print(f"\n--- KONTROL BAŞLATILDI ({time.strftime('%H:%M:%S')}) ---")
    
    # 1. Adım: Mailleri kontrol et ve sadece hepsi varsa indir
    success = check_and_download_specific_mails(senders, keywords)
    
    if success:
        # 2. Adım: İndirilen dosyaları işle
        print("Adım 2: Veriler ana dosyaya işleniyor...")
        process_excel_files()
        print("--- TÜM SÜREÇ BAŞARIYLA TAMAMLANDI ---\n")
        return True # İşlem bitti
    else:
        print("Eksik mail olduğu için işlem yapılamadı. Bir sonraki kontrolde tekrar denenecek.")
        return False # Hala bekliyoruz

def scheduled_job(senders, keywords):
    # Belirlenen saatte kontrolü başlat
    # Eğer eksik varsa 15 dakikada bir tekrar dene
    completed = run_bot_logic(senders, keywords)
    
    if not completed:
        # Geçici bir görev oluştur: 15 dakikada bir çalış ve tamamlanınca kendini iptal et
        def retry():
            nonlocal completed
            completed = run_bot_logic(senders, keywords)
            if completed:
                return schedule.CancelJob

        schedule.every(15).minutes.do(retry)

def main():
    print("--- ÖZEL MAİL VERİ ÇEKME BOTU ---")
    
    # Gönderenleri al
    sender_input = input("Mail beklediğiniz kişilerin e-posta adreslerini girin (Virgülle ayırın): ")
    senders = [s.strip() for s in sender_input.split(",") if s.strip()]
    
    if len(senders) < 2:
        print("Uyarı: En az 2 kişi girmelisiniz (X ve Y).")
        # Test için varsayılanlar eklenebilir veya kullanıcı zorlanabilir

    # Konu filtresi al
    subject_input = input("Maillerin konusunda geçmesi gereken ortak kelimeleri girin: ")
    keywords = [k.strip() for k in subject_input.split(",") if k.strip()]

    run_time = input("Botun her gün ilk kontrolü yapacağı saati girin (Örn: 09:30): ")
    
    # Zamanlayıcıyı kur
    schedule.every().day.at(run_time).do(scheduled_job, senders=senders, keywords=keywords)
    print(f"\nBot kuruldu! Her gün saat {run_time}'da başlayacak ve tüm mailler gelene kadar 15 dk'da bir deneyecek.")
    
    # Manuel çalıştırma
    ilk_calisma = input("Hemen şimdi kontrol edilsin mi? (e/h): ")
    if ilk_calisma.lower() == 'e':
        run_bot_logic(senders, keywords)

    while True:
        schedule.run_pending()
        time.sleep(30)

if __name__ == "__main__":
    main()
