import tkinter as tk
import time
import os
import json
import datetime

# Biến kiểm soát việc đang chờ hay thoát ứng dụng
app_should_exit = True

def gettime():
    now = datetime.datetime.now()
    time = f"{now.month}{now.day}{now.hour}"
    return time


def show_big_message_fullscreen(message):
    """Hiển thị bảng thông báo lớn, vẫn để lộ thanh task bar để người dùng có thể tắt máy.
       Bổ sung ô nhập thời gian & passcode 6 số.
       Trả về số phút gia hạn nếu bấm 'Chơi thêm', hoặc None nếu đóng thông báo/bấm Đã hiểu.
    """
    result = {"minutes": None}  # mutable để sửa ở closure

    root = tk.Tk()
    root.title("Thông báo quan trọng")

    # Đặt cửa sổ đủ lớn nhưng không fullscreen, vẫn để lộ thanh taskbar
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    window_width = screen_width
    window_height = int(screen_height * 0.96)  # Bớt một phần cho thanh taskbar

    x = 0
    y = 0

    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.attributes("-topmost", True)
    root.configure(bg='#f0f0f0')

    root.protocol("WM_DELETE_WINDOW", lambda: None)
    root.overrideredirect(True)

    label = tk.Label(
        root,
        text=message,
        font=("Helvetica", 50, "bold"),
        fg="green",
        bg="#f0f0f0",
        wraplength=int(window_width * 0.8),
        justify="center"
    )
    label.pack(expand=True)

    frame_inputs = tk.Frame(root, bg="#f0f0f0")
    frame_inputs.pack()

    # Thêm text area ở giữa màn hình để có thể gõ chữ, có scroll bar và luôn focus vào dòng cuối
    text_area_frame = tk.Frame(root, bg="#f0f0f0")
    text_area_frame.pack(pady=18)
    text_area_label = tk.Label(text_area_frame, text="Trao đổi:", font=("Helvetica", 15), bg="#f0f0f0")
    text_area_label.pack(anchor="w")
    text_area_scrollbar = tk.Scrollbar(text_area_frame)
    text_area_scrollbar.pack(side="right", fill="y")
    text_area = tk.Text(
        text_area_frame, 
        font=("Helvetica", 15), 
        height=5, 
        width=72, 
        wrap="word", 
        yscrollcommand=text_area_scrollbar.set
    )
    text_area.pack(side="left", fill="both", expand=True)
    text_area_scrollbar.config(command=text_area.yview)

    # Hàm luôn focus vào dòng cuối cùng của text area
    def focus_text_area_last_line(event=None):
        text_area.see("end")
        text_area.focus_set()
    # Gắn event mỗi khi nội dung thay đổi hoặc được focus, luôn cuộn xuống cuối
    text_area.bind("<<Modified>>", lambda e: (focus_text_area_last_line(), text_area.edit_modified(0)))
    text_area.bind("<FocusIn>", focus_text_area_last_line)

    # Ô nhập thời gian (có thể để trống hoặc nhập số phút)
    tk.Label(frame_inputs, text="Nhập thời gian (phút):", font=("Helvetica", 15), bg="#f0f0f0").grid(row=0, column=0, padx=4, pady=6, sticky="e")
    time_entry = tk.Entry(frame_inputs, font=("Helvetica", 15), width=8, justify="center")
    time_entry.grid(row=0, column=1, padx=4, pady=6)
    # Ô nhập passcode 6 số
    tk.Label(frame_inputs, text="Passcode 6 số:", font=("Helvetica", 15), bg="#f0f0f0").grid(row=1, column=0, padx=4, pady=6, sticky="e")
    pass_entry = tk.Entry(frame_inputs, font=("Helvetica", 15), width=8, show="*", justify="center")
    pass_entry.grid(row=1, column=1, padx=4, pady=6)
    pass_entry.focus_set()

    error_label = tk.Label(root, text="", fg="red", font=("Helvetica", 13), bg="#f0f0f0")
    error_label.pack(pady=(0,8))

    def on_close():
        # Kiểm tra passcode và thời gian
        entered_pass = pass_entry.get()
        entered_time = time_entry.get()
        times = gettime() #datetime.datetime.now().strftime("%m%d")
        if entered_pass != times:
            error_label.config(text="Sai passcode! Vui lòng thử lại.")
            return
        # Kiểm tra trường thời gian (cho phép rỗng hoặc số dương)
        if entered_time:
            try:
                t = int(entered_time)
                if t < 0:
                    error_label.config(text="Thời gian phải >= 0.")
                    return
            except ValueError:
                error_label.config(text="Thời gian phải là số nguyên.")
                return
        result["minutes"] = 18
        root.destroy()

    btn_close = tk.Button(
        root, text="Đã hiểu", font=("Helvetica", 13), command=on_close,
        bg="#4CAF50", fg="white", padx=14, pady=3
    )
    btn_close.place(relx=1.0, x=-14, y=11, anchor="ne")

    # Đổi lại flow: Nút 'Chơi thêm' sẽ đóng cửa sổ và trả về số phút, main kiểm soát logic
    def extend_time():
        passcode = pass_entry.get()
        times = gettime()
        minutes = time_entry.get()
        try:
            minutes = int(minutes)
            if minutes <= 0:
                raise ValueError
        except ValueError:
            error_label.config(text="Nhập số phút hợp lệ.")
            return

        if not passcode.isdigit():
            error_label.config(text="Passcode không hợp lệ.")
            return

        if passcode == times:
            error_label.config(text="")
            result["minutes"] = minutes
            root.destroy()
        else:
            error_label.config(text="Passcode sai. Thử lại!")

    btn_extend = tk.Button(
        root, text="Thêm giờ chơi", font=("Helvetica", 15), command=extend_time,
        bg="#2196F3", fg="white", padx=20
    )
    btn_extend.pack(pady=(0, 20))

    root.mainloop()
    return result["minutes"]

def show_big_message(message):
    """Hiển thị bảng thông báo (kiểu cũ), trả về số phút hoãn nếu có."""
    return show_big_message_fullscreen(message)

# --- Số lần chạy ứng dụng trong ngày ---
COUNTER_FILE = "reminder_count.json"
COUNTER_TIME_FILE = "reminder_time.txt"

def load_today_count():
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    if not os.path.exists(COUNTER_FILE):
        return 0, today_str
    try:
        with open(COUNTER_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("date") == today_str:
            return data.get("count", 0), today_str
        else:
            return 0, today_str
    except Exception:
        return 0, today_str

def save_today_count(count, today_str):
    try:
        with open(COUNTER_FILE, "w", encoding="utf-8") as f:
            json.dump({"date": today_str, "count": count}, f)
    except Exception:
        pass



# --- Chức năng delay khi người dùng chọn "Chơi thêm" ---
def wait_and_show_message(msg, seconds):
    """Hàm helper để thực hiện chờ và hiển thị message,
    có xử lý chơi thêm - hoãn nhiều lần liên tiếp nếu cần."""
    while True:
        print(f"Ứng dụng sẽ hiện thông báo sau {seconds} giây...")
        #time.sleep(seconds)
        # Đếm thời gian thực sự trôi qua kể từ lúc bắt đầu sleep cho tới khi bị tắt app bất thường
        start_time = time.time()
        elapsed = ELAPSED
        try:
            for i in range(seconds):
                time.sleep(1)
                elapsed = elapsed + 1
                if elapsed % 60 == 0 and elapsed != 0:
                    # Lưu lại thời gian đã qua mỗi 1 phút (ví dụ, ghi vào tệp)
                    with open(COUNTER_TIME_FILE, "w", encoding="utf-8") as f:
                        f.write(str(elapsed))
        except KeyboardInterrupt:
            pass
        
        Total_elapsed = int((time.time() - start_time + 59) // 60)  # phút, làm tròn lên   
        print(f"Đã trôi qua {Total_elapsed} phút kể từ khi bắt đầu chờ.")
        # Ghi log ra file mỗi lần hết giờ
        try:
            with open("reminder_log.txt", "a", encoding="utf-8") as log_file:
                log_file.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Lần chạy thứ {count} - Đã hết {Total_elapsed} phút\n")
        except Exception as e:
            print(f"Lỗi ghi log: {e}")
   
        res_minutes = show_big_message(msg)
        if res_minutes is None:
            # Đóng message: không chọn chơi thêm, thoát vòng lặp - app kết thúc
            break
        else:
            # Người dùng chọn chơi thêm, tiếp tục chờ số phút mới
            # (Đổi sang *60 cho đúng phút)
            seconds = int(res_minutes) * 60


def load_elaped():
    """Đọc giá trị elapsed từ file COUNTER_TIME_FILE nếu có"""
    try:
        with open(COUNTER_TIME_FILE, "r", encoding="utf-8") as f:
            content = f.read()
            return int(content.strip()) if content.strip().isdigit() else 0
    except Exception as e:
        print(f"Exception in load_elaped: {e}")
        return 0


import urllib.request
import hashlib
import shutil

CURRENT_FOLDER = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_FOLDER = CURRENT_FOLDER + "/download"
DEST_FOLDER = CURRENT_FOLDER

def file_checksum(filepath):
    """Tính checksum SHA-256 của file"""
    try:
        sha256 = hashlib.sha256()
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                sha256.update(chunk)
        return sha256.hexdigest()
    except Exception as e:
        print(f"Lỗi tính checksum cho {filepath}: {e}")
        return None

def download_reminder_pyw(url, download_folder=DOWNLOAD_FOLDER, dest_folder=DEST_FOLDER):
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)

    download_path = os.path.join(download_folder, "reminder.pyw")
    dest_path = os.path.join(dest_folder, "reminder.pyw")
    try:
        urllib.request.urlretrieve(url, download_path)
        print(f"Đã tải reminder.pyw về {download_path}")
    except Exception as e:
        print(f"Lỗi khi tải file: {e}")
        return False

    # So sánh checksum file mới tải và file ở dest_folder (nếu có)
    checksum_download = file_checksum(download_path)
    checksum_dest = file_checksum(dest_path) if os.path.exists(dest_path) else None

    if (checksum_dest is None) or (checksum_download != checksum_dest):
        try:
            shutil.copy2(download_path, dest_path)
            print(f"Đã thay thế file ở {dest_path} bằng file mới tải.")
        except Exception as e:
            print(f"Lỗi khi copy file sang dest: {e}")
            return False
    else:
        print("File đã giống nhau, không cần thay thế.")

    return True

# Ví dụ sử dụng:
url_online = "https://ala168.github.io/qlhv/reminder.pyw"
download_reminder_pyw(url_online)


# Khi mở app: kiểm tra và tăng số lần chạy, nếu đã từng chạy hôm nay thì hiện thông báo luôn
count, today = load_today_count()
if count == 0 & os.path.exists(COUNTER_TIME_FILE):
    try:
        os.remove(COUNTER_TIME_FILE)
    except Exception as e:
        print(f"Lỗi khi xóa file {COUNTER_TIME_FILE}: {e}")
count += 1
save_today_count(count, today)


WAIT_SECONDS = 20*60
ELAPSED = WAIT_SECONDS #load_elaped()
print(f"Đã mở {count-1} lần, lần trước đã sử dụng {ELAPSED} giây, còn lại {WAIT_SECONDS - ELAPSED} giây...")

if count > 1:
    # Gọi show_big_message để lấy phút hoãn từ entry_time, rồi chuyển thành giây
    msg = "HÔM NAY ĐÃ HẾT GIỜ CHƠI RỒI!\nHÃY TẮT MÁY VÀ ĐỨNG DẬY VẬN ĐỘNG ĐI THÔI"
    if ELAPSED < WAIT_SECONDS:
        WAIT_SECONDS = WAIT_SECONDS - ELAPSED
    else:
        res_minutes = show_big_message(msg)
        if res_minutes is not None:
            WAIT_SECONDS = int(res_minutes) * 60
        else:
            WAIT_SECONDS = 0
else:
    msg = "HẾT GIỜ DÙNG MÁY TÍNH RỒI!\nHÃY TẮT MÁY VÀ ĐỨNG DẬY VẬN ĐỘNG ĐI THÔI"

if WAIT_SECONDS > 0:
    wait_and_show_message(msg, WAIT_SECONDS)