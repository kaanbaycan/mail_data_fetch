import os
import win32com.client
from datetime import datetime

# --- AYARLAR ---
FILTER_SENDER = ""  # Örn: "ornek@sirket.com" (Boş bırakılırsa herkesi kabul eder)
FILTER_KEYWORDS = ["rapor", "excel", "veri"]  # Konuda geçmesi gereken kelimeler
SAVE_DIR = "indirilen_ekler"

def check_and_download_specific_mails(required_senders, keywords, folder_name="Inbox"):
    """
    Belirlenen klasör altında beklenen mailleri arar.
    """
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Klasörü bul
        root_folder = outlook.GetDefaultFolder(6) # Inbox
        target_folder = root_folder

        if folder_name.lower() != "inbox":
            try:
                # Önce Inbox'ın altında ara
                target_folder = root_folder.Folders[folder_name]
            except:
                try:
                    # Bulamazsa ana dizinde (Inbox ile aynı seviyede) ara
                    target_folder = root_folder.Parent.Folders[folder_name]
                except:
                    print(f"HATA: '{folder_name}' klasörü bulunamadı! 'Inbox' kullanılıyor.")
                    target_folder = root_folder

        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)

        found_senders = set()
        mails_to_download = []

        print(f"'{target_folder.Name}' klasörü taranıyor... (Hedef: {required_senders})")

        # Son 24 saatteki maillere bakmak mantıklı olabilir
        for message in messages:
            try:
                sender = message.SenderEmailAddress.lower()
                subject = message.Subject.lower()
                
                # Bu mail beklediğimiz kişilerden birinden mi?
                matching_sender = next((s for s in required_senders if s.lower() in sender), None)
                
                if matching_sender and any(kw.lower() in subject for kw in keywords):
                    if message.Attachments.Count > 0:
                        found_senders.add(matching_sender)
                        mails_to_download.append(message)
                
                # Eğer tüm gönderenleri bulduysak aramayı durdur (en güncelleri aldık)
                if len(found_senders) == len(required_senders):
                    break
            except:
                continue

        # KURAL: Her iki mail de gelmiş mi?
        if len(found_senders) == len(required_senders):
            print("BAŞARILI: Beklenen tüm mailler bulundu. İndirme başlıyor...")
            for msg in mails_to_download:
                for attachment in msg.Attachments:
                    if attachment.FileName.endswith((".xlsx", ".xls", ".csv")):
                        file_path = os.path.join(os.getcwd(), SAVE_DIR, attachment.FileName)
                        attachment.SaveAsFile(file_path)
                        print(f"İndirildi: {attachment.FileName} (Gönderen: {msg.SenderEmailAddress})")
            return True
        else:
            missing = set(required_senders) - found_senders
            print(f"EKSİK: Şu kişilerden mail bekleniyor: {missing}")
            return False

    except Exception as e:
        print(f"Hata: {e}")
        return False

if __name__ == "__main__":
    download_attachments()
