import time
import psutil
import os
import json
from datetime import datetime

#pip install psutil
#& "C:/Users/Abc/AppData/Local/Programs/Python/Python313/python.exe" -m pip install psutil 2>&1



def is_application_running(APPLICATION_NAME):
    """Kiểm tra xem APPLICATION_NAME có đang chạy không."""
    for proc in psutil.process_iter(attrs=['name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() == APPLICATION_NAME:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

USAGE_LOG_FILE = "app_usage_log.json"

def load_daily_usage():
    """
    Load usage log, return list of dicts with keys: date, app, seconds.
    Example: [{"date": "2024-06-10", "app": "chrome.exe", "seconds": 1234}]
    """
    if os.path.exists(USAGE_LOG_FILE):
        try:
            with open(USAGE_LOG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Lỗi khi đọc log: {e}")
    return []

def save_daily_usage(usage_data):
    """
    Save usage log as a list of dicts with keys: date, app, seconds.
    """
    try:
        with open(USAGE_LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(usage_data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Lỗi khi ghi log: {e}")

def find_today_entry(usage_data, today_str, app_name):
    """
    Tìm trong usage_data entry có ngày và tên app trùng, trả về index nếu có, ngược lại trả về -1
    """
    for idx, entry in enumerate(usage_data):
        if entry.get('date') == today_str and entry.get('app') == app_name:
            return idx
    return -1

def track_application_running_time(APPLICATION_NAME, poll_interval=5):
    """Theo dõi thời gian APPLICATION_NAME đang chạy, ghi lại tổng thời gian theo ngày, tên ứng dụng.
    Ghi log sau mỗi một phút."""
    print(f"Đang theo dõi tiến trình {APPLICATION_NAME}. Nhấn Ctrl+C để thoát.")
    running = False
    total_time = 0
    start_time = None

    usage_data = load_daily_usage()
    today_str = datetime.now().strftime('%Y-%m-%d')
    app_name = APPLICATION_NAME
    entry_idx = find_today_entry(usage_data, today_str, app_name)
    if entry_idx != -1:
        total_time = usage_data[entry_idx].get('seconds', 0)
    else:
        # Nếu chưa có entry, thêm entry mới
        usage_data.append({"date": today_str, "app": app_name, "seconds": 0})
        entry_idx = len(usage_data) - 1
        total_time = 0

    last_save_time = time.time()

    try:
        while True:
            app_running = is_application_running(app_name)
            current_time = time.time()
            if app_running and not running:
                # APP vừa mới khởi chạy
                start_time = current_time
                running = True
                print(f"{APPLICATION_NAME} đã khởi động.")
            elif not app_running and running:
                # APP vừa mới đóng
                session_time = current_time - start_time
                total_time += session_time
                print(f"{APPLICATION_NAME} vừa đóng. Thời gian chạy phiên này: {session_time:.2f} giây.")
                running = False
                start_time = None
                # Ghi log
                usage_data[entry_idx]['seconds'] = int(total_time)
                save_daily_usage(usage_data)
                last_save_time = current_time
            # Hiện thời gian tổng nếu APP đang chạy
            if running:
                elapsed = total_time + (current_time - start_time)
                print(f"Thời gian {APPLICATION_NAME} đang chạy hôm nay: {elapsed:.2f} giây.", end='\r')
                # Ghi log sau mỗi một phút chạy
                if current_time - last_save_time >= 60:
                    usage_data[entry_idx]['seconds'] = int(elapsed)
                    save_daily_usage(usage_data)
                    last_save_time = current_time
            else:
                print(f"Tổng thời gian {APPLICATION_NAME} đã chạy hôm nay: {total_time:.2f} giây.", end='\r')
            time.sleep(poll_interval)
    except KeyboardInterrupt:
        # Khi người dùng thoát chương trình
        if running:
            session_time = time.time() - start_time
            total_time += session_time
            print(f"\n{APPLICATION_NAME} đang chạy, cộng thêm phiên hiện tại: {session_time:.2f} giây.")
        print(f"\nTổng thời gian {APPLICATION_NAME} đã chạy hôm nay: {total_time:.2f} giây.")
        # Ghi log khi thoát
        usage_data[entry_idx]['seconds'] = int(total_time)
        save_daily_usage(usage_data)

if __name__ == "__main__":
    track_application_running_time("chrome.exe")

 