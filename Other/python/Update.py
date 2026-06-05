import datetime
from datetime import datetime, timedelta

def generate(offset_blocks: int = 0) -> str:
    secret_key=26
    #offset_blocks=0

    now = datetime.now()
    
    if offset_blocks != 0:
        now += timedelta(minutes=offset_blocks * 2)
        
    hour = now.hour
    minute = now.minute
    
    time_block = minute // 2
    
    first_part = (hour * 7 + time_block + secret_key) % 100
    
    second_part = (first_part * 3 + time_block * 5) % 100
    
    return f"{first_part:02d}{second_part:02d}"

# --- KIỂM TRA CHẠY THỬ ---
if __name__ == "__main__":
    print(f"Thời gian hiện tại: {datetime.now().strftime('%H:%M:%S')}")
    print(f"Mã hiện tại (Dùng trong 2 phút): {generate()}")
    print(f"Mã chu kỳ kế tiếp (+2 phút): {generate(offset_blocks=1)}")
