import os
import win32com.client
from datetime import datetime, date

SAVE_DIR = "indirilen_ekler"

def find_folder_recursive(folders, name):
    """Outlook klasörleri içinde ismi eşleşeni derinlemesine arar."""
    for folder in folders:
        if folder.Name.lower() == name.lower():
            return folder
        res = find_folder_recursive(folder.Folders, name)
        if res: return res
    return None

def check_and_download_specific_mails(folder_name="jet fuel"):
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
    else:
        for f in os.listdir(SAVE_DIR):
            try: os.remove(os.path.join(SAVE_DIR, f))
            except: pass

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Daha sağlam klasör bulma (Tüm hesapları tarar)
        target_folder = None
        for store in outlook.Stores:
            target_folder = find_folder_recursive(store.GetRootFolder().Folders, folder_name)
            if target_folder: break

        if not target_folder:
            print(f"HATA: '{folder_name}' klasörü hiçbir hesapta bulunamadı!")
            return False

        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)

        today = date.today()
        found_mails = []

        print(f"'{target_folder.Name}' taranıyor. Bugünün ekli mailleri aranıyor...")
        
        for message in messages:
            try:
                # Sadece bugünün mailleri
                if message.ReceivedTime.date() != today:
                    if message.ReceivedTime.date() < today: break
                    continue
                
                # SADECE EXCEL EKİ OLANLARI SAY
                has_excel = False
                for att in message.Attachments:
                    if att.FileName.lower().endswith((".xlsx", ".xls")):
                        has_excel = True
                        break
                
                if has_excel:
                    found_mails.append(message)
                    print(f"Uygun mail bulundu: {message.Subject}")
                
                if len(found_mails) == 2: break
            except: continue

        if len(found_mails) == 2:
            print("2 adet ekli mail bulundu. İndiriliyor...")
            for i, msg in enumerate(found_mails):
                suffix = "X_file.xlsx" if i == 0 else "Y_file.xlsx"
                for att in msg.Attachments:
                    if att.FileName.lower().endswith((".xlsx", ".xls")):
                        att.SaveAsFile(os.path.join(os.getcwd(), SAVE_DIR, suffix))
                        print(f" -> {suffix} kaydedildi.")
                        break
            return True
        else:
            print(f"HATA: Bugün sadece {len(found_mails)} adet Excel ekli mail bulundu.")
            return False

    except Exception as e:
        print(f"Outlook Hatası: {e}")
        return False
