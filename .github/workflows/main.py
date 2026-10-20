import time
from datetime import datetime

def main():
    print("启动无限循环任务...")
    count = 0
    
    while True:
        count += 1
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # ==========================================
        # 这里写你的核心业务逻辑
        # ==========================================
        print(f"[{current_time}] 第 {count} 次循环执行中...")
        
        # 每次循环休息 10 秒，防止把 CPU 跑满
        time.sleep(10) 

if __name__ == "__main__":
    main()
