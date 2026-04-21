import time
import psutil

#pip install psutil
#& "C:/Users/Abc/AppData/Local/Programs/Python/Python313/python.exe" -m pip install psutil 2>&1

APPLICATION_NAME = 'chrome.exe' #'notepad++.exe'

def is_application_running():
    """Kiểm tra xem APPLICATION_NAME có đang chạy không."""
    for proc in psutil.process_iter(attrs=['name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() == APPLICATION_NAME:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def track_application_running_time(poll_interval=5):
    """Theo dõi thời gian APPLICATION_NAME đang chạy."""
    print(f"Đang theo dõi tiến trình {APPLICATION_NAME}. Nhấn Ctrl+C để thoát.")
    running = False
    total_time = 0
    start_time = None
    try:
        while True:
            notepad_running = is_application_running()
            if notepad_running and not running:
                # APPLICATION_NAME vừa mới khởi chạy
                start_time = time.time()
                running = True
                print(f"{APPLICATION_NAME} đã khởi động.")
            elif not notepad_running and running:
                # APPLICATION_NAME vừa mới đóng
                session_time = time.time() - start_time
                total_time += session_time
                print(f"{APPLICATION_NAME} vừa đóng. Thời gian chạy phiên này: {session_time:.2f} giây.")
                running = False
                start_time = None
            # Hiện thời gian tổng nếu APPLICATION_NAME đang chạy
            if running:
                elapsed = total_time + (time.time() - start_time)
                print(f"Thời gian {APPLICATION_NAME} đang chạy: {elapsed:.2f} giây.", end='\r')
            else:
                print(f"Tổng thời gian {APPLICATION_NAME} đã chạy: {total_time:.2f} giây.", end='\r')
            time.sleep(poll_interval)
    except KeyboardInterrupt:
        # Khi người dùng thoát chương trình
        if running:
            session_time = time.time() - start_time
            total_time += session_time
            print(f"\n{APPLICATION_NAME} đang chạy, cộng thêm phiên hiện tại: {session_time:.2f} giây.")
        print(f"\nTổng thời gian {APPLICATION_NAME} đã chạy: {total_time:.2f} giây.")

if __name__ == "__main__":
    track_application_running_time()