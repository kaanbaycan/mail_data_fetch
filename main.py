from outlook_bot import download_attachments
from veri_isleyici import process_excel_files
import os

def main():
    print("--- BOT BAŞLATILDI ---")
    
    # 1. Adım: Mailleri kontrol et ve ekleri indir
    print("\nAdım 1: Outlook mailleri taranıyor...")
    download_attachments()
    
    # 2. Adım: İndirilen dosyaları işle
    print("\nAdım 2: Veriler ana dosyaya işleniyor...")
    process_excel_files()
    
    print("\n--- İŞLEM TAMAMLANDI ---")

if __name__ == "__main__":
    main()
