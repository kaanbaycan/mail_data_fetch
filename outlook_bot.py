import os
import win32com.client
from datetime import datetime, date

SAVE_DIR = "indirilen_ekler"

def check_and_download_specific_mails(folder_name="jet fuel"):
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        root_folder = outlook.GetDefaultFolder(6)
        
        try:
            target_folder = root_folder.Folders[folder_name]
        except:
            try:
                target_folder = root_folder.Parent.Folders[folder_name]
            except:
                print(f"HATA: '{folder_name}' klasörü bulunamadı!")
                return False

        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)

        today = date.today()
        found_mails = []

        for message in messages:
            try:
                if message.ReceivedTime.date() == today:
                    found_mails.append(message)
                if len(found_mails) == 2 or message.ReceivedTime.date() < today:
                    break
            except: continue

        if len(found_mails) == 2:
            print("\n2 mail bulundu. Gönderen ismine göre kaydediliyor...")
            for i, msg in enumerate(found_mails):
                # Gönderen ismini temizle (dosya adı olacağı için)
                s_name = "".join(x for x in msg.SenderName if x.isalnum())
                for attachment in msg.Attachments:
                    if attachment.FileName.endswith((".xlsx", ".xls")):
                        # Dosya adının başına gönderen ismini ekle
                        file_name = f"{s_name}_{attachment.FileName}"
                        file_path = os.path.join(os.getcwd(), SAVE_DIR, file_name)
                        attachment.SaveAsFile(file_path)
                        print(f" -> Kaydedildi: {file_name}")
            return True
        else:
            print(f"Eksik mail: Bugün şu ana kadar {len(found_mails)} mail geldi.")
            return False
    except Exception as e:
        print(f"Hata: {e}")
        return False
