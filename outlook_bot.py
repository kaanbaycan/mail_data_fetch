import os
import win32com.client
from datetime import datetime

# --- AYARLAR ---
SAVE_DIR = "indirilen_ekler"

def get_sender_email(message):
    """Outlook mesajından temiz SMTP adresini almaya çalışır."""
    try:
        # Önce standart adresi dene
        email = message.SenderEmailAddress
        # Eğer Exchange formatındaysa (/o=...) SMTP adresini çek
        if email and "/o=" in email:
            try:
                if message.SenderEmailType == "EX":
                    return message.Sender.GetExchangeUser().PrimarySmtpAddress
            except:
                pass
        return email
    except:
        return "Bilinmiyor"

def check_and_download_specific_mails(required_senders, keywords, folder_name="Inbox"):
    """
    Belirlenen klasördeki mailleri listeler ve ekleri indirir.
    """
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Klasör bulma
        root_folder = outlook.GetDefaultFolder(6)
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
        messages.Sort("[ReceivedTime]", True)

        print(f"\n--- '{target_folder.Name}' Klasörü İçeriği (Son 10 Mail) ---")
        
        found_count = 0
        for i, message in enumerate(messages):
            if i >= 10: break
            
            try:
                subject = message.Subject
                sender_email = get_sender_email(message)
                sender_name = message.SenderName
                
                print(f"[{i+1}] Konu: {subject}")
                print(f"    Gönderen: {sender_name} <{sender_email}>")

                # Şimdilik filtre yapmadan TÜM EXCELLERİ indiriyoruz
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        if attachment.FileName.endswith((".xlsx", ".xls", ".csv")):
                            file_path = os.path.join(os.getcwd(), SAVE_DIR, attachment.FileName)
                            attachment.SaveAsFile(file_path)
                            print(f"    -> EK İNDİRİLDİ: {attachment.FileName}")
                            found_count += 1
            except Exception as e:
                continue

        print(f"\nToplam {found_count} adet ek indirildi.")
        return True if found_count > 0 else False

    except Exception as e:
        print(f"Hata: {e}")
        return False

if __name__ == "__main__":
    # Test amaçlı
    check_and_download_specific_mails([], [], "jet fuel")
