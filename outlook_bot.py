import os
import win32com.client
from datetime import datetime, date

SAVE_DIR = "indirilen_ekler"

def get_sender_email(message):
    try:
        email = message.SenderEmailAddress
        if email and "/o=" in email:
            try:
                if message.SenderEmailType == "EX":
                    return message.Sender.GetExchangeUser().PrimarySmtpAddress
            except: pass
        return email.lower() if email else ""
    except: return ""

def check_and_download_specific_mails(sender_x, sender_y, folder_name="jet fuel"):
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
        found_x = False
        found_y = False

        print(f"\n--- Bugünün Mailleri Taranıyor ({sender_x} ve {sender_y}) ---")

        for message in messages:
            if message.ReceivedTime.date() != today:
                if message.ReceivedTime.date() < today: break
                continue

            email = get_sender_email(message)
            
            # X kişisi mi?
            if not found_x and sender_x.lower() in email:
                for attachment in message.Attachments:
                    if attachment.FileName.endswith((".xlsx", ".xls")):
                        attachment.SaveAsFile(os.path.join(os.getcwd(), SAVE_DIR, "X_file.xlsx"))
                        print(f" -> X (BCSL) dosyası indirildi.")
                        found_x = True
                        break
            
            # Y kişisi mi?
            elif not found_y and sender_y.lower() in email:
                for attachment in message.Attachments:
                    if attachment.FileName.endswith((".xlsx", ".xls")):
                        attachment.SaveAsFile(os.path.join(os.getcwd(), SAVE_DIR, "Y_file.xlsx"))
                        print(f" -> Y (AAZBN) dosyası indirildi.")
                        found_y = True
                        break

            if found_x and found_y: break

        return found_x and found_y
    except Exception as e:
        print(f"Hata: {e}")
        return False
