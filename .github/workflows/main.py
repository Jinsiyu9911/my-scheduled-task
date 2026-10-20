#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电表电量监控 - 定时任务入口（GitHub Actions 等平台用）
版本: v4.1  修改时间: 2026-10-20

职责边界：
    - main.py    只负责「跑多久、每轮间隔、什么时候安全退出、退出时怎么交代清楚」
    - monitor.py 保持独立可运行（python monitor.py / --loop / --chart ...），
                 自身不感知平台；这里只是把它当成一个模块来调用。

文件布局（注意：三个文件都放在 .github/workflows/ 下）：
    .github/workflows/main.py      ← 本文件，定时任务入口
    .github/workflows/monitor.py   ← 业务代码，独立可运行
    .github/workflows/main.yml     ← workflow 定义
    因此 main.py 里用「脚本自身所在目录」定位 monitor.py，
    而 monitor.py 的 PROJECT_DIR 也等于这个目录 → data/ 会落在 .github/workflows/data/。
    workflow 里必须相应提交 .github/workflows/data，否则 runner 上的数据会随虚拟机一起消失。

每轮流程：
    import monitor → monitor.setup_console()（UTF-8）→ monitor.setup_log()（写 data/monitor.log）
    → monitor.run_once()（抓取 → CSV → 告警 → 整点出图）
    每轮都用新子进程 + _MP_MAIN 标记，避免每轮重复打开日志文件、模块级状态残留，
    以及任何一处 sys.exit 把整个进程带走。

平台注意（GitHub Actions）：
    - 单 job 最长 6 小时会被强杀，这里留 15 分钟余量主动 exit(0) 收尾；
    - Actions 的 runner 是 Ubuntu/Windows 虚拟机，无交互桌面：
      monitor.notify() 会失败，告警正文会落到 data/alerts.log（已有兜底逻辑）；
      真要在 Actions 里推送通知，把 --webhook 指向企业微信/钉钉/飞书机器人即可。
    - data/ 里的 CSV、图表、alert_state.json 必须靠 workflow 里的 commit 或 artifacts 保住，
      否则 14 天保留期也只剩一次运行的记忆。workflow 通过「进程结束时扫 data/ 有无变化」
      决定要不要提交：有变化才 commit（避免 6 小时窗口里狂推 300 次），
      并在退出后延时几秒再 commit，让 runner 上的 git 写完、日志落地。
      数据目录是「与 main.py 同目录」的 data/，即 .github/workflows/data/。
"""

import argparse
import json
import os
import subprocess
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

# ====================== 配置 ======================
MAX_RUN_TIME = 20700            # 5小时45分钟 = 20700 秒；6 小时硬上限前留 15 分钟余量
SLEEP_SECONDS = 60              # 每轮间隔：与 monitor.py 的「每分钟采集」节奏一致
WAIT_ALIGNED = True             # 是否对齐到整分再跑（monitor.py 的实时图表靠 minute == 0 触发）
PROGRESS_EVERY = 30             # 每 N 轮打一次进度（避免 6 小时刷出 300 多行一样的日志）
WEBHOOK_URL = ""                # 可选：企业微信/钉钉/飞书机器人的 webhook 地址（留空=不发）

# 边跑边提交：每 30 分钟把 data/ 里的小文件提交并推回仓库，0 表示关闭。
# 只提交 CSV/JSON 这类小文件（几 KB/次）；charts/*.bmp 每张近 1MB，
# 所以图表先转成 PNG 再提交（见 PUBLISH_CHARTS），BMP 本身只留在 artifact 里。
AUTO_COMMIT_EVERY = 60 # 1800
COMMIT_FILES = ("meter_data.csv", "latest.json", "alert_state.json",
                ".last_prune", ".ready_to_commit", ".commit_status")

# 图表发布：monitor.py 只会写 BMP（每张约 938KB），直接提交进 git 一个月能涨几百 MB。
# 这里用标准库 zlib 把它转成同名 PNG（实测 3~7KB）再提交，成本降到 1/150，
# 于是「整点生成图表」也能在 1 分钟内出现在仓库里。
PUBLISH_CHARTS = True

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据目录：默认就在 main.py 旁边（.github/workflows/data），但 GitHub 明令禁止
# GITHUB_TOKEN 提交或修改 `.github/workflows/**` 下的任何文件，会报
#   refusing to allow a GitHub App to create or update workflow `.../.last_prune`
#   without `workflows` permission
# 而 `permissions:` 里没有 workflows 这个键，无法授权。所以 CI 下用 METER_DATA_DIR
# 把数据指到仓库根的 data/（main.yml 里已设置），代码留在 .github/workflows/ 不动。
# 相对路径按「启动时的当前目录」解析成绝对路径：子进程的 cwd 是 .github/workflows，
# 若保留相对路径，monitor.py 会把它解析到 .github/workflows/data 去，又触发同一限制。
DATA_ROOT = os.path.abspath(
    os.environ.get("METER_DATA_DIR") or os.path.join(PROJECT_DIR, "data"))


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
        latest_path = os.path.join(DATA_ROOT, "latest.json")
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


# ====================== 图表转 PNG（纯标准库） ======================
def _bmp_to_png(bmp_path, png_path):
    """把 monitor.py 生成的未压缩 24 位 BMP 转成 PNG。

    不引入 Pillow：PNG 的结构简单到用 zlib + struct 就够——IHDR/IDAT/IEND 三个块，
    每行前加一个 filter 字节 0，像素从 BMP 的 BGR 换序成 PNG 的 RGB。
    实测 800x400 的图表：空图 3.4KB、24 根柱子的满图 6.4KB，约 BMP 的 1/150。
    """
    import struct
    import zlib

    with open(bmp_path, "rb") as f:
        data = f.read()
    if data[:2] != b"BM":
        raise ValueError("不是 BMP 文件")
    off = struct.unpack_from("<I", data, 10)[0]
    w, h = struct.unpack_from("<ii", data, 18)
    bpp = struct.unpack_from("<H", data, 28)[0]
    compression = struct.unpack_from("<I", data, 30)[0]
    if bpp != 24 or compression != 0 or h <= 0 or w <= 0:
        raise ValueError(f"只支持未压缩 24 位下→上 BMP (bpp={bpp}, comp={compression}, h={h})")

    stride = (w * 3 + 3) // 4 * 4
    raw = bytearray()
    for y in range(h):
        start = off + (h - 1 - y) * stride        # BMP 像素是从下往上存的
        line = data[start:start + w * 3]
        inter = bytearray(w * 3)
        inter[0::3] = line[2::3]                  # BGR -> RGB，用切片让 C 层做搬砖
        inter[1::3] = line[1::3]
        inter[2::3] = line[0::3]
        raw.append(0)                             # PNG 每行前缀：filter type 0 (None)
        raw.extend(inter)

    def chunk(tag, payload):
        return (struct.pack(">I", len(payload)) + tag + payload +
                struct.pack(">I", zlib.crc32(tag + payload) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)   # 8bit / truecolor RGB
    png = (b"\x89PNG\r\n\x1a\n" +
           chunk(b"IHDR", ihdr) +
           chunk(b"IDAT", zlib.compress(bytes(raw), 9)) +
           chunk(b"IEND", b""))
    with open(png_path, "wb") as f:
        f.write(png)


def publish_charts(verbose=True):
    """把 charts/ 下新生成的 BMP 转成 PNG（PNG 才会被提交进仓库）。返回转换个数。"""
    if not PUBLISH_CHARTS:
        return 0
    chart_dir = os.path.join(DATA_ROOT, "charts")
    if not os.path.isdir(chart_dir):
        return 0
    converted = 0
    for name in sorted(os.listdir(chart_dir)):
        if not name.lower().endswith(".bmp"):
            continue
        bmp = os.path.join(chart_dir, name)
        png = bmp[:-4] + ".png"
        try:
            if os.path.exists(png) and os.path.getmtime(png) >= os.path.getmtime(bmp):
                continue
            _bmp_to_png(bmp, png)
            converted += 1
            if verbose:
                print(f"[chart] BMP -> PNG: {os.path.basename(png)} "
                      f"({os.path.getsize(png) // 1024}KB)")
        except Exception as e:
            print(f"[chart] 转换失败 {name}: {type(e).__name__}: {e}")
    return converted


# ====================== 数据提交（边跑边推） ======================
REPO_ROOT_CACHE = []            # 缓存 git 仓库根目录；空列表=还没查过，[None]=不是 git 仓库


def _git(args, cwd):
    try:
        p = subprocess.run(["git"] + args, cwd=cwd, capture_output=True,
                           text=True, timeout=60)
    except (OSError, subprocess.SubprocessError) as e:
        return -1, "", f"{type(e).__name__}: {e}"
    return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()


def repo_root():
    """返回 git 仓库根目录；不在仓库里（比如本地直接跑）返回 None。"""
    if REPO_ROOT_CACHE:
        return REPO_ROOT_CACHE[0]
    code, out, _ = _git(["rev-parse", "--show-toplevel"], PROJECT_DIR)
    root = out if code == 0 and out else None
    if root:
        root = os.path.abspath(root)
    REPO_ROOT_CACHE.append(root)
    return root


def _push_token():
    """取用于推送的 token。

    GITHUB_TOKEN 在「Settings → Actions → Workflow permissions」为只读时无法推送，
    而且 YAML 里的 permissions 只能降权不能提权。所以优先使用 PAT（GH_PUSH_TOKEN /
    PAT_TOKEN 环境变量），它本来就带着 contents:write。
    """
    for name in ("GH_PUSH_TOKEN", "PAT_TOKEN", "GITHUB_TOKEN"):
        token = os.environ.get(name) or ""
        if token.strip():
            return token.strip()
    return ""


def _remote_url(root):
    code, out, _ = _git(["remote", "get-url", "origin"], root)
    return out if code == 0 else ""


def _authed_url(root, token):
    """把 origin 的 https 地址加上 token（仅用于本次 push/pull，不写进 git config）。"""
    url = _remote_url(root)
    if not token or not url.startswith("https://") or "github.com" not in url:
        return ""
    return url.replace("https://", f"https://x-access-token:{token}@", 1)


def _clean(text, token):
    """把命令输出里的 token 抹掉，防止泄漏到公开的 Actions 日志。"""
    if not text:
        return ""
    if token:
        text = text.replace(token, "***")
    return text


def _write_commit_status(ok, detail):
    """把最近一次提交结果写进 data/.commit_status（会被一起提交，仓库里直接可查）。"""
    try:
        path = os.path.join(DATA_ROOT, ".commit_status")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       "ok": ok, "detail": detail[:400]}, f, ensure_ascii=False)
    except OSError:
        pass


def commit_data(verbose=True):
    """把 data/ 里的小文件提交并推回当前分支。返回 True 表示真的提交了。

    设计取舍：
      - 只提交 COMMIT_FILES（CSV/JSON，每次增量几 KB）与 charts/*.png，
        不提交 charts/*.bmp（每张约 938KB）与 monitor.log，否则 git 历史会迅速膨胀；
      - 推送失败（没权限、网络抖动）只记日志并写 .commit_status，绝不打断主循环；
      - 不在 git 仓库里（本地开发）直接跳过。
    """
    root = repo_root()
    if not root:
        if verbose:
            print("[commit] 跳过：当前目录不在 git 仓库内")
        return False

    data_rel = os.path.relpath(DATA_ROOT, root).replace("\\", "/")
    data_abs = DATA_ROOT
    targets = [f"{data_rel}/{name}" for name in COMMIT_FILES
               if os.path.exists(os.path.join(data_abs, name))]
    # 图表只提交转换后的 PNG（体积约 BMP 的 1/50）
    chart_dir = os.path.join(data_abs, "charts")
    if os.path.isdir(chart_dir):
        targets += [f"{data_rel}/charts/{name}" for name in sorted(os.listdir(chart_dir))
                    if name.lower().endswith(".png")]
    if not targets:
        if verbose:
            print("[commit] 跳过：data/ 里没有可提交的文件")
        return False

    token = _push_token()

    code, _, err = _git(["add", "-A", "--"] + targets, root)
    if code != 0:
        err = _clean(err, token)
        print(f"[commit] git add 失败: {err}")
        _write_commit_status(False, f"git add 失败: {err}")
        return False

    # Actions 上默认没有 git 身份，缺失时补一个（已配置过就不动用户自己的设置）
    if _git(["config", "user.name"], root)[0] != 0:
        _git(["config", "user.name", "meter-monitor-bot"], root)
    if _git(["config", "user.email"], root)[0] != 0:
        _git(["config", "user.email", "meter-monitor-bot@users.noreply.github.com"], root)

    code, _, _ = _git(["diff", "--cached", "--quiet"], root)
    if code == 0:                 # 没有变化
        if verbose:
            print("[commit] 跳过：data/ 无变化")
        return False

    msg = f"data: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 电表数据（自动提交）"
    code, _, err = _git(["commit", "-q", "-m", msg], root)
    if code != 0:
        err = _clean(err, token)
        print(f"[commit] git commit 失败: {err}")
        _write_commit_status(False, f"git commit 失败: {err}")
        return False

    _, sha, _ = _git(["rev-parse", "--short", "HEAD"], root)
    print(f"[commit] 已提交 {sha}: {msg}")

    # 推送：接力期间可能有别的运行刚推过，先 rebase 再推，重试 3 次
    branch = os.environ.get("GITHUB_REF_NAME") or "main"
    authed = _authed_url(root, token)
    target = authed or "origin"
    for attempt in range(1, 4):
        if authed:
            _git(["pull", "--rebase", "--autostash", target, branch], root)
        else:
            _git(["pull", "--rebase", "--autostash", "origin", branch], root)
        code, _, err = _git(["push", target, f"HEAD:{branch}"], root)
        if code == 0:
            print(f"[commit] 已推送 (第 {attempt} 次尝试) -> origin/{branch}")
            _write_commit_status(True, f"已推送 {sha}")
            return True
        err = _clean(err, token)
        print(f"[commit] 推送失败（第 {attempt} 次）: {err[:300]}")
        time.sleep(2)

    err = _clean(err, token)
    print("[commit] 推送多次失败，数据先留在 runner 上；本轮继续跑")
    print("[commit] 常见原因：Settings → Actions → General → Workflow permissions "
          "未设为 Read and write，或 PAT 缺少 contents:write")
    _write_commit_status(False, f"推送失败: {err[:300]}")
    return True


def _maybe_commit(last_commit_ts):
    """到点就提交一次，返回新的时间戳。"""
    if AUTO_COMMIT_EVERY <= 0:
        return last_commit_ts
    if time.time() - last_commit_ts < AUTO_COMMIT_EVERY:
        return last_commit_ts
    commit_data(verbose=False)
    return time.time()


# ====================== 单轮执行 ======================
def run_once():
    """在子进程里跑一轮真正的采集。返回 (退出码, 是否触发过告警, 耗秒)。

    用新的 Python 子进程 + _MP_MAIN 环境变量标记：
      - monitor.run_once() 内部抓取/存盘失败会 sys.exit(1|2)，子进程隔离后不会带走本循环；
      - 每轮一个干净进程，日志句柄、urllib 连接、进程级状态都不残留。
    """
    env = dict(os.environ)
    env["_MP_MAIN"] = "1"           # 防止 monitor.py 若将来加了 __main__ 守卫以外的副作用
    env["METER_DATA_DIR"] = DATA_ROOT   # 让 monitor.py 把数据写到 DATA_ROOT（CI 下是仓库根 data/）
    cmd = [sys.executable, os.path.join(PROJECT_DIR, "monitor.py")]
    t0 = time.time()
    proc = subprocess.run(cmd, cwd=PROJECT_DIR, env=env)
    elapsed = time.time() - t0
    return proc.returncode, False, elapsed


def run_once_inline():
    """同一进程内调用 monitor.run_once()（性能更好，但一轮失败会抛 SystemExit）。"""
    import monitor
    os.environ.setdefault("METER_DATA_DIR", DATA_ROOT)
    monitor.DATA_DIR = Path(DATA_ROOT)
    monitor.CHART_DIR = Path(DATA_ROOT) / "charts"
    monitor.DATA_FILE = Path(DATA_ROOT) / "meter_data.csv"
    monitor.ALERT_FILE = Path(DATA_ROOT) / "alert_state.json"
    monitor.LATEST_FILE = Path(DATA_ROOT) / "latest.json"
    monitor.ALERTS_LOG = Path(DATA_ROOT) / "alerts.log"
    monitor.LOG_FILE = Path(DATA_ROOT) / "monitor.log"
    monitor.PRUNE_MARKER = Path(DATA_ROOT) / ".last_prune"
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


# ====================== 接力退出标记 ======================
READY_MARKER = os.path.join(DATA_ROOT, ".ready_to_commit")


def _mark_ready_for_commit(delay=3.0):
    """写一个 READY 标记后停几秒。

    workflow 在「本步骤跑完」之后才做 git 提交（见 main.yml），这里先刷盘 + 留出缓冲，
    顺便把标记文件也放进 data/，让 data/ 目录永远非空，避免 git 提交步骤出错。
    """
    try:
        os.makedirs(os.path.dirname(READY_MARKER), exist_ok=True)
        with open(READY_MARKER, "w", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    except OSError as e:
        print(f"[main] 写 READY 标记失败: {e}")
    for stream_ in (sys.stdout, sys.stderr):
        try:
            stream_.flush()
        except Exception:
            pass
    time.sleep(delay)


# ====================== 主循环 ======================
def main(argv=None):
    global MAX_RUN_TIME, SLEEP_SECONDS, WAIT_ALIGNED, AUTO_COMMIT_EVERY

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
    p.add_argument("--commit-every", type=int, default=AUTO_COMMIT_EVERY,
                   help=f"每 N 秒把 data/ 小文件提交并推回仓库，默认 {AUTO_COMMIT_EVERY}；0=关闭")
    args = p.parse_args(argv)

    MAX_RUN_TIME = args.max_run
    SLEEP_SECONDS = max(1, args.interval)
    WAIT_ALIGNED = not args.no_align
    AUTO_COMMIT_EVERY = max(0, args.commit_every)

    setup_console()
    start_time = time.time()
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("任务启动...")
    print(f"  python      = {sys.version.split()[0]} ({sys.executable})")
    print(f"  工作目录     = {PROJECT_DIR}")
    print(f"  最长运行     = {MAX_RUN_TIME}s ({MAX_RUN_TIME / 3600:.2f}h，平台 6h 上限前主动退出)")
    print(f"  轮间隔       = {SLEEP_SECONDS}s，对齐整分 = {WAIT_ALIGNED}")
    print(f"  子进程模式   = {not args.inline}")
    print(f"  自动提交     = {'关闭' if AUTO_COMMIT_EVERY <= 0 else f'每 {AUTO_COMMIT_EVERY}s'} "
          f"(git 仓库: {repo_root() or '不在仓库内，跳过'})")
    print(f"  启动时间     = {started_at}")

    count = 0
    failures = 0
    last_commit = 0          # 0 => 第一轮就提交一次，不用等到第一个间隔
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
            _mark_ready_for_commit(delay=0)
            publish_charts()
            commit_data()
            return code

        # 图表整点生成后，立刻转成 PNG 备提交（只有整点那一轮会真的转换）
        publish_charts(verbose=False)

        # 边跑边提交：到点就把 data/ 里的小文件推回仓库，
        # 这样网页上随时能看到 CSV，也不用怕 6 小时中途被取消导致数据全丢
        last_commit = _maybe_commit(last_commit)

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
            # 退出前再提交一次，把最后这段时间的数据也推上去
            _mark_ready_for_commit(delay=0)
            publish_charts(verbose=False)
            commit_data(verbose=False)
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
