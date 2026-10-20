import time
import sys
from datetime import datetime

def main():
    print("任务启动...")
    start_time = time.time()
    
    # 设定运行时间：5小时45分钟 = 20700 秒。
    # (留出15分钟余量，确保在GitHub 6小时强制拔电源之前，我们能自己安全退出)
    MAX_RUN_TIME = 70 #20700 

    count = 0
    while True:
        count += 1
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # ==========================================
        # 在这里写入你真正的业务逻辑
        # ==========================================
        print(f"[{current_time}] 第 {count} 次循环执行中...")
        
        # 模拟每次任务耗时
        time.sleep(10) 
        # ==========================================

        # 核心：检查是否快到达 6 小时极限
        elapsed_time = time.time() - start_time
        if elapsed_time > MAX_RUN_TIME:
            print(f"已运行 {elapsed_time} 秒，即将达到平台限制。")
            print("主动保存状态，正常退出，准备触发下一轮接力...")
            # sys.exit(0) 表示正常结束，GitHub 会判定这一步为 success
            sys.exit(0) 

if __name__ == "__main__":
    main()
