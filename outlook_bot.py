import os
import win32com.client
from datetime import datetime

# --- AYARLAR ---
FILTER_SENDER = ""  # Örn: "ornek@sirket.com" (Boş bırakılırsa herkesi kabul eder)
FILTER_KEYWORDS = ["rapor", "excel", "veri"]  # Konuda geçmesi gereken kelimeler
SAVE_DIR = "indirilen_ekler"

def download_attachments():
    # Klasör yoksa oluştur
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
        print(f"Klasör oluşturuldu: {SAVE_DIR}")

    try:
        # Outlook bağlantısı
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Gelen Kutusu (6 = olFolderInbox)
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        
        # Mesajları tarihe göre sırala (Sondan başa)
        messages.Sort("[ReceivedTime]", True)

        print("Mailler kontrol ediliyor...")

        for message in messages:
            try:
                subject = message.Subject
                sender = message.SenderEmailAddress
                
                # Filtreleme Kontrolleri
                sender_match = not FILTER_SENDER or (FILTER_SENDER.lower() in sender.lower())
                keyword_match = any(kw.lower() in subject.lower() for kw in FILTER_KEYWORDS)

                if sender_match and keyword_match:
                    if message.Attachments.Count > 0:
                        print(f"Uygun mail bulundu: {subject}")
                        for attachment in message.Attachments:
                            if attachment.FileName.endswith((".xlsx", ".xls", ".csv")):
                                file_path = os.path.join(os.getcwd(), SAVE_DIR, attachment.FileName)
                                attachment.SaveAsFile(file_path)
                                print(f"Ek indirildi: {attachment.FileName}")
                        
                        # Sadece en güncel olanı indirmek istersen burada break diyebilirsin
                        # break 

            except Exception as e:
                # Bazı sistem mesajları hataya sebep olabilir, pas geçiyoruz
                continue

    except Exception as e:
        print(f"Hata oluştu: {e}")
        print("Not: Bu script sadece Windows üzerinde Outlook masaüstü uygulaması yüklüyken çalışır.")

if __name__ == "__main__":
    download_attachments()
