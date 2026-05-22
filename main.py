from outlook_bot import download_attachments
from veri_isleyici import process_excel_files
import schedule
import time
import os

def run_bot(keywords):
    print(f"\n--- İŞLEM BAŞLATILDI ({time.strftime('%H:%M:%S')}) ---")
    
    # 1. Adım: Mailleri kontrol et ve ekleri indir
    print(f"Adım 1: Outlook mailleri taranıyor... (Aranan: {keywords})")
    download_attachments(keywords=keywords)
    
    # 2. Adım: İndirilen dosyaları işle
    print("Adım 2: Veriler ana dosyaya işleniyor...")
    process_excel_files()
    
    print("--- İŞLEM TAMAMLANDI ---\n")
    print("Sıradaki çalışma saati bekleniyor... (Durdurmak için CTRL+C)")

def main():
    print("--- MAİL VERİ ÇEKME BOTU ZAMANLAYICI ---")
    
    # Konu filtresi al
    subject_input = input("Maillerin konusunda geçmesi gereken kelimeleri girin (Virgülle ayırın, örn: rapor, günlük, satis): ")
    keywords = [k.strip() for k in subject_input.split(",") if k.strip()]
    
    if not keywords:
        keywords = ["rapor", "excel", "veri"]
        print(f"Varsayılan filtreler kullanılacak: {keywords}")

    run_time = input("Botun her gün çalışmasını istediğiniz saati girin (Örn: 09:30): ")
    
    try:
        # Zamanlayıcıyı kur
        schedule.every().day.at(run_time).do(run_bot, keywords=keywords)
        print(f"\nBot kuruldu! Her gün saat {run_time} olduğunda '{', '.join(keywords)}' kelimelerini arayacak.")
        
        # İlk başta bir kere çalıştırılsın mı?
        ilk_calisma = input("Hemen şimdi bir kere çalıştırılsın mı? (e/h): ")
        if ilk_calisma.lower() == 'e':
            run_bot(keywords)

        while True:
            schedule.run_pending()
            time.sleep(60) # Her dakika kontrol et
            
    except Exception as e:
        print(f"Hata: Saat formatı yanlış olabilir. Lütfen 'SS:DD' formatında girin. ({e})")

if __name__ == "__main__":
    main()
