import os
import win32com.client
from datetime import datetime, date

SAVE_DIR = "indirilen_ekler"

def check_and_download_specific_mails(folder_name="jet fuel"):
    """
    Belirlenen klasörde BUGÜN gelen son 2 mailin eklerini indirir.
    """
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
    else:
        # Önce klasörü temizle (Eski dosyalar kalmasın)
        for f in os.listdir(SAVE_DIR):
            try: os.remove(os.path.join(SAVE_DIR, f))
            except: pass
        print(f"Klasör temizlendi: {SAVE_DIR}")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        root_folder = outlook.GetDefaultFolder(6) # Inbox
        
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
                if message.ReceivedTime.date() == today:
                    found_mails.append(message)
                
                if len(found_mails) == 2:
                    break
                
                if message.ReceivedTime.date() < today:
                    break
            except:
                continue

        if len(found_mails) == 2:
            print("\nBAŞARILI: Bugün gelen 2 mail bulundu. Ekler indiriliyor...")
            # İlkini X_file, ikincisini Y_file olarak kaydet (İşleyici için)
            for i, msg in enumerate(found_mails):
                suffix = "X_file.xlsx" if i == 0 else "Y_file.xlsx"
                for attachment in msg.Attachments:
                    if attachment.FileName.endswith((".xlsx", ".xls")):
                        file_path = os.path.join(os.getcwd(), SAVE_DIR, suffix)
                        attachment.SaveAsFile(file_path)
                        print(f"    -> {suffix} olarak kaydedildi.")
            return True
        else:
            print(f"\nEKSİK: Bugün sadece {len(found_mails)} mail bulundu. 2 mail bekleniyor.")
            return False

    except Exception as e:
        print(f"Hata: {e}")
        return False
