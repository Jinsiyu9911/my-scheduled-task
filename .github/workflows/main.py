#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电表电量监控 - 定时任务入口（GitHub Actions 等平台用）
版本: v4.1  修改时间: 2026-10-20

职责边界：
    - main.py    只负责「跑多久、每轮间隔、什么时候安全退出、退出时怎么交代清楚」
    - monitor.py 保持独立可运行（python monitor.py / --loop / --chart ...），
                 自身不感知平台；这里只是把它当成一个模块来调用。

每轮流程：
    import monitor → monitor.setup_console()（UTF-8）→ monitor.setup_log()（写 data/monitor.log）
    → monitor.run_once()（抓取 → CSV → 告警 → 整点出图）
    每轮都用新子进程 + _MP_MAIN 标记，避免每轮重复打开日志文件、模块级状态残留，
    以及任何一处 sys.exit 把整个进程带走。

平台注意（GitHub Actions）：
    - 单 job 最长 6 小时会被强杀，这里留 15 分钟余量主动 exit(0) 收尾；
    - Actions 的 runner 是 Ubuntu/Windows 虚拟机，无交互桌面：
      monitor.notify() 会失败，告警正文会落到 data/alerts.log（已有兜底逻辑）；
      真要在 Actions 里推送通知，把 --notify-webhook 指向企业微信/钉钉/飞书机器人即可。
    - data/ 里的 CSV、图表、alert_state.json 必须靠 workflow 里的 commit 或 artifacts 保住，
      否则 14 天保留期也只剩一次运行的记忆。
"""

import argparse
import json
import os
import sys
import time
import urllib.request
from datetime import datetime

# ====================== 配置 ======================
MAX_RUN_TIME = 20700            # 5小时45分钟 = 20700 秒；6 小时硬上限前留 15 分钟余量
SLEEP_SECONDS = 60              # 每轮间隔：与 monitor.py 的「每分钟采集」节奏一致
WAIT_ALIGNED = True             # 是否对齐到整分再跑（monitor.py 的实时图表靠 minute == 0 触发）
PROGRESS_EVERY = 30             # 每 N 轮打一次进度（避免 6 小时刷出 300 多行一样的日志）
WEBHOOK_URL = ""                # 可选：企业微信/钉钉/飞书机器人的 webhook 地址（留空=不发）

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))


# ====================== 工具函数 ======================
def setup_console():
    """统一 UTF-8 输出，避免中文在 GBK 控制台 / Actions 日志里乱码。"""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None:
                stream.reconfigure(encoding="utf-8", errors="replace",
                                   line_buffering=True)
        except Exception:
            pass


def sleep_until_next_minute():
    """睡到下一个整分，保证 monitor.py 里「整点那一分钟」生成图表的条件能被命中。"""
    now = time.time()
    time.sleep(max(1.0, SLEEP_SECONDS - (now % SLEEP_SECONDS)))


def fetch_webhook_proxy(url):
    """可选通知兜底：把 data/latest.json 的内容推给机器人 webhook。

    平台 runner 没有桌面，系统 Toast 不可能弹出来；如果配置了 webhook，
    就在告警/汇总时把同一条消息推出去（webhook 服务通常有频率限制，这里只推告警类）。
    """
    try:
        latest_path = os.path.join(PROJECT_DIR, "data", "latest.json")
        with open(latest_path, "r", encoding="utf-8") as f:
            latest = json.load(f)
    except OSError:
        return False
    text = (f"电表监控 {latest.get('timestamp')} | 剩余 {latest.get('kwh')} kWh"
            f" | 状态 {latest.get('status')}")
    payload = json.dumps({"msgtype": "text", "text": {"content": text}}).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers={
        "Content-Type": "application/json; charset=utf-8"})
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            resp.read()
        return True
    except Exception as e:
        print(f"[webhook] 推送失败: {type(e).__name__}: {e}")
        return False


# ====================== 单轮执行 ======================
def run_once():
    """在子进程里跑一轮真正的采集。返回 (退出码, 是否触发过告警, 耗秒)。

    用新的 Python 子进程 + _MP_MAIN 环境变量标记：
      - monitor.run_once() 内部抓取/存盘失败会 sys.exit(1|2)，子进程隔离后不会带走本循环；
      - 每轮一个干净进程，日志句柄、urllib 连接、进程级状态都不残留。
    """
    import subprocess

    env = dict(os.environ)
    env["_MP_MAIN"] = "1"           # 防止 monitor.py 若将来加了 __main__ 守卫以外的副作用
    cmd = [sys.executable, os.path.join(PROJECT_DIR, "monitor.py")]
    t0 = time.time()
    proc = subprocess.run(cmd, cwd=PROJECT_DIR, env=env)
    elapsed = time.time() - t0
    return proc.returncode, False, elapsed


def run_once_inline():
    """同一进程内调用 monitor.run_once()（性能更好，但一轮失败会抛 SystemExit）。"""
    import monitor
    monitor.setup_console()
    monitor.setup_log()
    try:
        monitor.run_once()
        return 0
    except SystemExit as e:
        return int(e.code or 0)
    finally:
        # run_once 内部已经把日志刷新到文件，这里不关闭，交给进程退出
        pass


# ====================== 主循环 ======================
def main(argv=None):
    global MAX_RUN_TIME, SLEEP_SECONDS, WAIT_ALIGNED

    p = argparse.ArgumentParser(description="电表监控定时任务入口（长跑 + 主动接力退出）")
    p.add_argument("--max-run", type=int, default=MAX_RUN_TIME,
                   help=f"最长运行秒数，默认 {MAX_RUN_TIME}（5h45m）")
    p.add_argument("--interval", type=int, default=SLEEP_SECONDS,
                   help=f"每轮间隔秒数，默认 {SLEEP_SECONDS}")
    p.add_argument("--inline", action="store_true",
                   help="在同一进程内直接调用 monitor.run_once()，不起子进程")
    p.add_argument("--once", action="store_true", help="只跑一轮后退出（调试用）")
    p.add_argument("--no-align", action="store_true", help="不等待对齐整分，逐轮立即跑")
    p.add_argument("--webhook", default=WEBHOOK_URL,
                   help="可选：告警栏 robot webhook 地址（企业微信/钉钉/飞书）")
    args = p.parse_args(argv)

    MAX_RUN_TIME = args.max_run
    SLEEP_SECONDS = max(1, args.interval)
    WAIT_ALIGNED = not args.no_align

    setup_console()
    start_time = time.time()
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("任务启动...")
    print(f"  python      = {sys.version.split()[0]} ({sys.executable})")
    print(f"  工作目录     = {PROJECT_DIR}")
    print(f"  最长运行     = {MAX_RUN_TIME}s ({MAX_RUN_TIME / 3600:.2f}h，平台 6h 上限前主动退出)")
    print(f"  轮间隔       = {SLEEP_SECONDS}s，对齐整分 = {WAIT_ALIGNED}")
    print(f"  子进程模式   = {not args.inline}")
    print(f"  启动时间     = {started_at}")

    count = 0
    failures = 0
    while True:
        count += 1
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n[{current_time}] ── 第 {count} 次循环 ──")

        # ==========================================
        # 真正的业务逻辑：交给 monitor.py
        # ==========================================
        if args.inline:
            code = 0
            try:
                code = run_once_inline()
            except SystemExit as e:
                code = int(e.code or 0)
            except Exception as e:
                code = 99
                print(f"[main] 本轮异常: {type(e).__name__}: {e}")
            elapsed = 0.0
        else:
            code, _, elapsed = run_once()

        if code == 0:
            print(f"[main] 本轮完成 (exit=0, {elapsed:.1f}s)")
        else:
            failures += 1
            print(f"[main] 本轮失败 (exit={code}, {elapsed:.1f}s)"
                  f" 累计失败 {failures} 次；{SLEEP_SECONDS}s 后重试")

        if args.webhook and code != 0:
            fetch_webhook_proxy(args.webhook)

        if args.once:
            print("[main] --once：单轮模式，退出")
            return code

        # 进度心跳：每 PROGRESS_EVERY 轮报一次运行时长，方便在 Actions 日志里定位
        if count % PROGRESS_EVERY == 0:
            up = time.time() - start_time
            print(f"[main] 进度：已跑 {count} 轮，{up / 60:.1f} 分钟，"
                  f"剩余额度 {max(0, MAX_RUN_TIME - up) / 60:.1f} 分钟")

        # 核心：检查是否快到达 6 小时极限
        elapsed_time = time.time() - start_time
        if elapsed_time > MAX_RUN_TIME:
            print(f"已运行 {elapsed_time:.0f} 秒，即将达到平台限制。")
            print(f"统计：共 {count} 轮，失败 {failures} 轮")
            print("主动保存状态，正常退出，准备触发下一轮接力...")
            # exit(0) 表示正常结束，平台会判定这一步为 success
            sys.exit(0)

        if WAIT_ALIGNED:
            sleep_until_next_minute()
        else:
            time.sleep(SLEEP_SECONDS)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[main] 收到 Ctrl+C，退出")
        sys.exit(130)
