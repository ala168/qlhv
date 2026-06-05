from datetime import datetime, timedelta

def generate() -> list:    
    valid_codes = []
    now = datetime.now()
    secret_key=26

    for offset in [-1, 0, 1]:
        check_time = now + timedelta(minutes=offset)
        hour = check_time.hour
        minute = check_time.minute
        code_hour = (hour*3 + secret_key) % 100
        code_minute = (minute + secret_key) % 100
        passcode = f"{code_hour:02d}{code_minute:02d}"
        valid_codes.append(passcode)
    return valid_codes

# --- KIỂM TRA THỬ NGHIỆM ---
if __name__ == "__main__":
    print(f"Thời gian hiện tại: {datetime.now().strftime('%H:%M:%S')}")
    codes = generate()
    print(f"Các passcode hợp lệ [-1p, Hiện tại, +1p]: {codes}")
    # Giả lập logic xác thực khi người dùng nhập mã
    user_input = input("Nhập passcode của bạn: ")
    if user_input not in generate():
        print("❌ Mã sai hoặc đã hết hạn!")
    else:
        print("🔓 Xác thực thành công!")
        