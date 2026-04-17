import urllib.request
import os

PUBLIC_REMINDER_DEST_FOLDER = "Q:/tiepvk/python/download"

def download_reminder_pyw(url, dest_folder=PUBLIC_REMINDER_DEST_FOLDER):
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    dest_path = os.path.join(dest_folder, "reminder.pyw")
    try:
        urllib.request.urlretrieve(url, dest_path)
        print(f"Đã tải reminder.pyw về {dest_path}")
        return True
    except Exception as e:
        print(f"Lỗi khi tải file: {e}")
        return False

# Ví dụ sử dụng:
url_online = "https://ala168.github.io/qlhv/reminder.pyw"
download_reminder_pyw(url_online)