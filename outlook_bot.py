import os
import win32com.client
from datetime import datetime, date

# --- AYARLAR ---
SAVE_DIR = "indirilen_ekler"

def check_and_download_specific_mails(required_senders, keywords, folder_name="Inbox"):
    """
    Belirlenen klasörde BUGÜN gelen son 2 mailin eklerini indirir.
    """
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Klasör bulma
        root_folder = outlook.GetDefaultFolder(6) # Inbox
        target_folder = root_folder

        if folder_name.lower() != "inbox":
            try:
                target_folder = root_folder.Folders[folder_name]
            except:
                try:
                    target_folder = root_folder.Parent.Folders[folder_name]
                except:
                    print(f"HATA: '{folder_name}' klasörü bulunamadı!")
                    return False

        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True) # En yeni en üstte

        today = date.today()
        found_mails = []

        print(f"\n--- '{target_folder.Name}' Klasöründe Bugünün Mailleri Taranıyor ---")
        
        for message in messages:
            try:
                # Mailin alındığı tarih
                received_date = message.ReceivedTime.date()
                
                if received_date == today:
                    found_mails.append(message)
                    print(f"Uygun mail bulundu: {message.Subject} ({message.ReceivedTime.strftime('%H:%M')})")
                
                # Sadece son 2 maili bulduysak dur
                if len(found_mails) == 2:
                    break
                
                # Eğer tarihler düne geçtiyse daha fazla bakmaya gerek yok (Sort olduğu için)
                if received_date < today:
                    break
            except:
                continue

        if len(found_mails) == 2:
            print("\nBAŞARILI: Bugün gelen 2 mail de bulundu. Ekler indiriliyor...")
            for msg in found_mails:
                for attachment in msg.Attachments:
                    if attachment.FileName.endswith((".xlsx", ".xls", ".csv")):
                        file_path = os.path.join(os.getcwd(), SAVE_DIR, attachment.FileName)
                        attachment.SaveAsFile(file_path)
                        print(f"    -> İndirildi: {attachment.FileName}")
            return True
        else:
            print(f"\nEKSİK: Bugün sadece {len(found_mails)} mail bulundu. 2 mail olması bekleniyor.")
            return False

    except Exception as e:
        print(f"Hata: {e}")
        return False
