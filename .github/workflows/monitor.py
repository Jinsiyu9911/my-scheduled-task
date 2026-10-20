#!/usr/bin/env python3
r"""
电表电量监控 - Windows PC 版
版本: v4.0  修改时间: 2026-10-20（由 Termux 安卓版 v3.5 移植）

- 纯标准库、零第三方依赖。通知由系统自带的 Windows PowerShell 5.1 + WinRT Toast 实现
  （Termux 的 termux-notification 已移除，也不需要安装 BurntToast 之类的模块）。
- 抓取电表页面，解析剩余电量与表具状态，追加写入 CSV。
- 告警推送到 Windows 通知中心：低电量、离线、恢复在线、每日汇总；同类型告警有冷却期。
- 图表为手写 24 位 BMP（不依赖 PIL/matplotlib）。实时图表只在「整点那一分钟」生成，
  每日汇总时再额外存一份快照。
- 数据保留期 2 周：CSV 只留最近 RETENTION_DAYS(=14) 天的记录，charts\ 只留
  RETENTION_DAYS 天内的快照，过期内容每天自动清理一次（所以实际最多多留不到 1 天）。

用法:
    python  monitor.py                # 单次采集（交给「任务计划程序」每分钟触发）
    pythonw monitor.py --loop         # 常驻，按整分钟对齐循环（推荐，免配任务计划程序）
    python  monitor.py --chart        # 只生成当前 24 小时图表，调试用
    python  monitor.py --test-notify  # 发一条测试 Toast，验证通知链路
    python  monitor.py --prune        # 立刻按保留期清理，不等每日自动清理

注册为登录自启:
    schtasks /create /tn "电表监控" /sc ONLOGON /f ^
             /tr "pythonw C:\code\meter_monitor\monitor.py --loop"

输出:
    控制台输出同时追加到 data\monitor.log（超过 2MB 自动滚动为 monitor.log.1）；
    Toast 发送失败时告警正文另外落到 data\alerts.log。
    CSV 为 UTF-8 with BOM，Excel 直接双击打开中文不乱码。
"""

import argparse
import base64
import csv
import json
import os
import re
import struct
import subprocess
import sys
import time
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

# ====================== 配置 ======================
METER_NO = "220406021323"
URL = f"http://api.app.znglxt.cn/site/index?meterNo={METER_NO}"
THRESHOLD_KWH = 100.0           # 低于此值告警
ALERT_COOLDOWN_HOURS = 6        # 同类型告警冷却时间（小时）
DAILY_SUMMARY_HOUR = 22         # 每日汇总时间：小时（24小时制）
DAILY_SUMMARY_MINUTE = 0        # 每日汇总时间：分钟
RETENTION_DAYS = 14             # 数据保留天数：CSV 记录 + 图表快照，最多留 2 周

# 通知 ID —— 在 Windows 下用作 Toast 的 Tag（Tag 上限 16 字符）
NID_LOW = "meter-low-kwh"
NID_OFFLINE = "meter-offline"
NID_SUMMARY = "meter-summary"
NID_TEST = "meter-test"

# Toast 使用的 AppUserModelID：未打包的普通脚本需要一个「已注册」的 AUMID，
# 通知才会进通知中心。默认借用 Windows PowerShell 开始菜单快捷方式的 AUMID，
# 失败再兜底尝试 Shell 的 AUMID；也可用环境变量 METER_TOAST_AUMID 覆盖。
TOAST_GROUP = "meter-monitor"
TOAST_APP_IDS = [
    os.environ.get("METER_TOAST_AUMID", "").strip(),
    r"{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe",
    "Microsoft.Windows.Shell.RunDialog",
]
TOAST_APP_IDS = [a for a in TOAST_APP_IDS if a]
NOTIFY_TIMEOUT = 20             # 单次 PowerShell 调用超时（秒）

# 路径（默认放在脚本所在目录的 data/ 下，与 Termux 版一致）
# 可用环境变量 METER_DATA_DIR 指定别的数据目录：GitHub Actions 禁止 GITHUB_TOKEN
# 提交/修改 `.github/workflows/**` 下的任何文件，所以 CI 里要把数据放到仓库其他位置。
PROJECT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("METER_DATA_DIR") or (PROJECT_DIR / "data"))
CHART_DIR = DATA_DIR / "charts"
DATA_FILE = DATA_DIR / "meter_data.csv"     # UTF-8 with BOM：Excel 双击打开中文不乱码
ALERT_FILE = DATA_DIR / "alert_state.json"
LATEST_FILE = DATA_DIR / "latest.json"
ALERTS_LOG = DATA_DIR / "alerts.log"
LOG_FILE = DATA_DIR / "monitor.log"
PRUNE_MARKER = DATA_DIR / ".last_prune"
LOG_MAX_BYTES = 2 * 1024 * 1024


# ====================== 控制台 / 日志 ======================
class _Tee:
    """把标准输出同时写进日志文件（任务计划程序、pythonw 下都没有可见控制台）。"""

    def __init__(self, stream, log_path):
        self._stream = stream
        self._fh = open(log_path, "a", encoding="utf-8")

    def write(self, s):
        if self._stream is not None:
            try:
                self._stream.write(s)
            except Exception:
                pass
        try:
            self._fh.write(s)
            self._fh.flush()
        except Exception:
            pass
        return len(s)

    def flush(self):
        for obj in (self._stream, self._fh):
            try:
                if obj is not None:
                    obj.flush()
            except Exception:
                pass

    def isatty(self):
        return False

    def __getattr__(self, name):
        return getattr(self._stream, name)


def setup_console():
    """统一使用 UTF-8 输出（避免中文在 GBK 控制台/管道里乱码）。无文件副作用，
    放在参数解析之前调用，连 --help 也不会乱码。"""
    if os.name == "nt":
        try:  # 控制台代码页切到 65001，配合下面的 reconfigure 才能正确显示中文
            import ctypes
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None:
                # line_buffering：输出被重定向到管道/文件时也能实时看到进度
                stream.reconfigure(encoding="utf-8", errors="replace",
                                   line_buffering=True)
        except Exception:
            pass


def setup_log():
    """把输出另外追加一份到 data/monitor.log（超过 2MB 滚动为 monitor.log.1）。"""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        if LOG_FILE.exists() and LOG_FILE.stat().st_size > LOG_MAX_BYTES:
            os.replace(LOG_FILE, LOG_FILE.with_name(LOG_FILE.name + ".1"))
        sys.stdout = _Tee(sys.stdout, LOG_FILE)
        sys.stderr = _Tee(sys.stderr, LOG_FILE)
    except OSError as e:
        print(f"[WARN] 日志文件初始化失败，仅输出到控制台: {e}")


# ====================== 通知（Windows Toast） ======================
# 注意：必须调用 powershell.exe（Windows PowerShell 5.1）。
# PowerShell 7(pwsh) 无法用 ContentType=WindowsRuntime 加载 WinRT 类型。
_TOAST_PS = r"""
$ErrorActionPreference = 'Stop'
try {
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime] | Out-Null
    $xmlText = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('@XML_B64@'))
    $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
    $doc.LoadXml($xmlText)
    $toast = New-Object Windows.UI.Notifications.ToastNotification -ArgumentList $doc
    $toast.Tag = '@TAG@'
    $toast.Group = '@GROUP@'
    $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('@APPID@')
    $notifier.Show($toast)
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
"""

_REMOVE_PS = r"""
$ErrorActionPreference = 'Stop'
try {
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    [Windows.UI.Notifications.ToastNotificationManager]::History.Remove('@APPID@', '@TAG@', '@GROUP@')
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
"""


def _xml_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _toast_xml(title, content, sticky):
    """拼 Toast XML。sticky 对应 Termux 版的 --ongoing（常驻），这里用 scenario=reminder
    让通知停在屏幕上直到手动关掉。"""
    scenario = ' scenario="reminder"' if sticky else ""
    return (
        f'<toast{scenario}>'
        '<visual><binding template="ToastGeneric">'
        f'<text>{_xml_escape(title)}</text>'
        f'<text>{_xml_escape(content)}</text>'
        '</binding></visual>'
        '</toast>'
    )


def _run_ps(script: str):
    """执行一段 Windows PowerShell 脚本，返回 (是否成功, 错误信息)。

    脚本用 -EncodedCommand 传递（UTF-16LE + Base64），彻底绕开引号、中文、换行的转义问题。
    """
    if os.name != "nt":
        return False, "当前不是 Windows 平台"
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    cmd = ["powershell.exe", "-NoProfile", "-NonInteractive",
           "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded]
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=NOTIFY_TIMEOUT,
                              creationflags=flags)
    except (OSError, subprocess.SubprocessError) as e:
        return False, f"{type(e).__name__}: {e}"
    if proc.returncode == 0:
        return True, ""
    raw = proc.stderr or proc.stdout or b""
    err = " ".join(raw.decode("utf-8", "replace").split())
    return False, (err[:300] or f"powershell 退出码 {proc.returncode}")


def _append_alert_log(title, content, err):
    """Toast 发不出去时，至少把告警正文记到文件里。"""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(ALERTS_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {title}\n{content}\n  通知失败: {err}\n\n")
    except OSError:
        pass


def notify(title, content, nid, ongoing=False, priority="high"):
    """把告警推送到 Windows 通知中心（Toast）。

    ongoing=True 对应 Termux 版的 --ongoing（常驻通知），这里映射为 scenario="reminder"。
    priority 仅保留旧调用签名，Toast 没有等价的优先级设置。
    返回 True 表示已投递；失败时告警正文写入 data/alerts.log，并沿用 Termux 版语义：
    本轮不更新告警状态，下一轮自动重试。
    """
    xml = _toast_xml(title, content, sticky=ongoing)
    xml_b64 = base64.b64encode(xml.encode("utf-8")).decode("ascii")

    errors = []
    for app_id in TOAST_APP_IDS:
        script = (_TOAST_PS
                  .replace("@XML_B64@", xml_b64)
                  .replace("@TAG@", nid)
                  .replace("@GROUP@", TOAST_GROUP)
                  .replace("@APPID@", app_id))
        ok, err = _run_ps(script)
        if ok:
            return True
        errors.append(f"{app_id}: {err}")

    _append_alert_log(title, content, " | ".join(errors))
    print(f"    [通知] Toast 发送失败: {errors[0] if errors else '未知原因'}")
    return False


def remove_notification(nid):
    """尽力撤掉指定 Tag 的 Toast（对应 Termux 版的 termux-notification-remove）。

    实测（Win10 19041）：History.Remove 返回成功但通知中心里的条目不会真的消失，
    所以这里只当「尽力而为」。真正让旧通知不见的是后续用同一个 Tag 再发一条——
    同 Tag + 同 Group 会自动覆盖旧 Toast，而本脚本两个调用点后面都紧跟同 Tag 的新通知。
    """
    if os.name != "nt":
        return
    for app_id in TOAST_APP_IDS:
        script = (_REMOVE_PS
                  .replace("@TAG@", nid)
                  .replace("@GROUP@", TOAST_GROUP)
                  .replace("@APPID@", app_id))
        ok, _ = _run_ps(script)
        if ok:
            return


# ====================== 数据抓取 ======================
def fetch_meter_data():
    """请求电表页面，返回 (剩余kWh, 状态)。"""
    req = urllib.request.Request(URL, headers={
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0.0.0 Safari/537.36'
        )
    })
    with urllib.request.urlopen(req, timeout=15) as resp:
        html = resp.read().decode('utf-8', errors='replace')

    kwh = None
    for pattern in [
        r'<span class="f-green">([\d.]+)\s*kWh</span>',
        r'>([\d.]+)\s*kWh<',
        r'剩余电量.*?>([\d.]+)',
    ]:
        m = re.search(pattern, html)
        if m:
            kwh = float(m.group(1))
            break

    status = "未知"
    for pattern in [
        r'<span class="f-green">(在线|离线)</span>',
        r'表具状态.*?(在线|离线)',
    ]:
        m = re.search(pattern, html)
        if m:
            status = m.group(1)
            break

    return kwh, status


# ====================== 数据存储 ======================
def save_data(kwh, status):
    """追加一条记录到 CSV，同时更新 latest.json."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    file_exists = DATA_FILE.exists()
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    mode = 'a' if file_exists else 'w'
    # 新建文件时用 utf-8-sig 写入 BOM（Excel 才不会把「在线」显示成乱码）；
    # 追加时用普通 utf-8，避免 BOM 出现在文件中间。读取端统一用 utf-8-sig，两种都兼容。
    encoding = 'utf-8' if file_exists else 'utf-8-sig'
    with open(DATA_FILE, mode, newline='', encoding=encoding) as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['timestamp', 'kwh', 'status'])
        writer.writerow([ts, str(kwh) if kwh is not None else "", status])

    latest = {
        'timestamp': ts,
        'kwh': kwh,
        'status': status,
        'meter_no': METER_NO,
    }
    LATEST_FILE.write_text(json.dumps(latest, ensure_ascii=False), encoding='utf-8')


def load_alert_state():
    if ALERT_FILE.exists():
        try:
            return json.loads(ALERT_FILE.read_text(encoding='utf-8'))
        except Exception:
            return {}
    return {}


def save_alert_state(state):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ALERT_FILE.write_text(
        json.dumps(state, indent=2, ensure_ascii=False), encoding='utf-8')


# ====================== 保留期清理（最多留 2 周） ======================
def _prune_csv():
    """把 CSV 裁剪到最近 RETENTION_DAYS 天，返回删除的行数。"""
    if not DATA_FILE.exists():
        return 0
    cutoff = datetime.now() - timedelta(days=RETENTION_DAYS)
    fieldnames = ['timestamp', 'kwh', 'status']
    kept = []
    removed = 0
    try:
        with open(DATA_FILE, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames or fieldnames
            for row in reader:
                try:
                    t = datetime.strptime(str(row.get('timestamp', '')),
                                          '%Y-%m-%d %H:%M:%S')
                except (ValueError, TypeError):
                    removed += 1        # 时间戳坏掉的行一并丢掉
                    continue
                if t >= cutoff:
                    kept.append(row)
                else:
                    removed += 1
    except OSError as e:
        print(f"  [清理] CSV 读取失败: {e}")
        return 0

    if removed == 0:
        return 0

    tmp = DATA_FILE.with_name(DATA_FILE.name + '.tmp')
    try:
        with open(tmp, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in kept:
                writer.writerow({k: row.get(k, '') for k in fieldnames})
        os.replace(tmp, DATA_FILE)      # 原子替换，避免读到写了一半的 CSV
    except OSError as e:
        print(f"  [清理] CSV 重写失败: {e}")
        try:
            tmp.unlink()
        except OSError:
            pass
        return 0
    return removed


def _prune_charts():
    """删掉 charts/ 下超过 RETENTION_DAYS 天的图表快照，返回删除个数。

    只处理 hourly_chart_<后缀>.bmp 这种快照；hourly_chart.bmp 是实时图，每次覆盖，不动。
    """
    if not CHART_DIR.is_dir():
        return 0
    cutoff_ts = (datetime.now() - timedelta(days=RETENTION_DAYS)).timestamp()
    removed = 0
    for path in CHART_DIR.glob('hourly_chart_*.bmp'):
        try:
            if path.stat().st_mtime < cutoff_ts:
                path.unlink()
                removed += 1
        except OSError:
            continue
    return removed


def prune_storage(force=False):
    """按保留期清理 CSV 与图表。默认每天只做一次（用 data/.last_prune 记录日期）。"""
    today = datetime.now().strftime('%Y-%m-%d')
    if not force:
        try:
            if PRUNE_MARKER.read_text(encoding='utf-8').strip() == today:
                return
        except OSError:
            pass

    removed_rows = _prune_csv()
    removed_files = _prune_charts()

    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        PRUNE_MARKER.write_text(today, encoding='utf-8')
    except OSError:
        pass

    if removed_rows or removed_files:
        print(f"  [清理] 保留 {RETENTION_DAYS} 天：CSV 删除 {removed_rows} 行，"
              f"图表删除 {removed_files} 个")


# ====================== 告警逻辑 ======================
def check_alerts(kwh, status):
    """检查是否需要发送通知。返回: [(类型, 标题, 内容, nid, ongoing), ...]"""
    alerts = []
    state = load_alert_state()
    now = datetime.now()

    def in_cooldown(key):
        last_str = state.get(key)
        if not last_str:
            return False
        try:
            last_time = datetime.fromisoformat(last_str)
            return (now - last_time).total_seconds() < ALERT_COOLDOWN_HOURS * 3600
        except Exception:
            return False

    # --- 低电量 ---
    if kwh is not None and kwh < THRESHOLD_KWH and not in_cooldown('low_kwh'):
        alerts.append((
            'low_kwh', '电量不足',
            f'剩余: {kwh} kWh | 阈值: {THRESHOLD_KWH} kWh\n表号: {METER_NO}\n请及时充值!',
            NID_LOW, True
        ))

    # --- 离线 ---
    if status == "离线" and not in_cooldown('offline'):
        alerts.append((
            'offline', '电表离线',
            f'表号: {METER_NO}\n状态: 离线\n请检查电表!',
            NID_OFFLINE, True
        ))

    # --- 恢复在线 ---
    # 只有真的推送过「离线」告警（state 里留着 offline 键）才会补发这条恢复通知
    if status == "在线" and state.get('offline'):
        remove_notification(NID_OFFLINE)
        alerts.append((
            'recovery', '电表恢复在线',
            f'表号: {METER_NO}\n当前电量: {kwh} kWh',
            NID_OFFLINE, False
        ))

    # --- 电量恢复正常 ---
    if kwh is not None and kwh >= THRESHOLD_KWH and state.get('low_kwh'):
        remove_notification(NID_LOW)
        alerts.append((
            'kwh_recovered', '电量恢复正常',
            f'当前: {kwh} kWh (阈值: {THRESHOLD_KWH})',
            NID_LOW, False
        ))

    # --- 每日汇总 ---
    # 当天 DAILY_SUMMARY_HOUR:MINUTE 之后首次运行触发，state 里记日期保证当天只发一次；
    # 如果 22:00~24:00 之间一次都没运行，当天的汇总不会补发（跨过午夜就不算当天了）。
    today_key = now.strftime('%Y-%m-%d')
    now_minutes = now.hour * 60 + now.minute
    trigger_minutes = DAILY_SUMMARY_HOUR * 60 + DAILY_SUMMARY_MINUTE

    if now_minutes >= trigger_minutes:
        if (state.get('daily_summary_date') != today_key and kwh is not None):
            consumption_24h = _calc_24h_consumption()
            summary = (
                f'[{today_key}] 电表 {METER_NO}\n'
                f'剩余电量: {kwh} kWh\n'
                f'状态: {status}'
            )
            if consumption_24h is not None:
                summary += f'\n近24h消耗: {consumption_24h:.2f} kWh'
                if consumption_24h > 0:
                    days_left = kwh / (consumption_24h)
                    summary += f'\n预估可用: {days_left:.0f} 天'
            alerts.append(('daily_summary', f'电量汇总 ({today_key})', summary, NID_SUMMARY, False))
            print(f"  [汇总] 已触发，通知栏推送详情")
        elif kwh is None:
            print(f"  [汇总] 跳过 — 电量数据为空")
        else:
            print(f"  [汇总] 跳过 — 今日已发 (date={state.get('daily_summary_date')})")
    else:
        print(f"  [汇总] 当前 {now.strftime('%H:%M')}, 触发时间 "
              f"{DAILY_SUMMARY_HOUR:02d}:{DAILY_SUMMARY_MINUTE:02d}, 尚未到点")

    return alerts


def update_alert_state(alert_types, status):
    state = load_alert_state()
    now = datetime.now().isoformat()
    for atype in alert_types:
        if atype == 'low_kwh':
            state['low_kwh'] = now
        elif atype == 'offline':
            state['offline'] = now
        elif atype == 'recovery':
            state.pop('offline', None)
        elif atype == 'kwh_recovered':
            state.pop('low_kwh', None)
        elif atype == 'daily_summary':
            state['daily_summary_date'] = datetime.now().strftime('%Y-%m-%d')
    state['last_status'] = status
    state['last_check'] = now
    save_alert_state(state)


def _calc_24h_consumption():
    """用 CSV 里最近 24h 的首末电量差估算消耗（依赖 CSV 按时间顺序追加）。"""
    if not DATA_FILE.exists():
        return None
    try:
        with open(DATA_FILE, 'r', encoding='utf-8-sig') as f:
            rows = list(csv.DictReader(f))
        if len(rows) < 2:
            return None
        now = datetime.now()
        first_kwh = None
        last_kwh = None
        for r in rows:
            try:
                t = datetime.strptime(r['timestamp'], '%Y-%m-%d %H:%M:%S')
                k = float(r['kwh'])
                if (now - t).total_seconds() <= 86400:
                    if first_kwh is None:
                        first_kwh = k
                    last_kwh = k
            except (ValueError, KeyError):
                continue
        if first_kwh is not None and last_kwh is not None:
            return first_kwh - last_kwh
        return None
    except Exception:
        return None


# ====================== 图表生成 ======================
def generate_hourly_chart(suffix: str = ""):
    """从 CSV 读取数据，生成每小时用电量 BMP 柱状图（最近 24 小时整点区间）。

    时间窗：起点 = 终点 - 24h；终点在有 suffix（每日快照）时取当天
    DAILY_SUMMARY_HOUR:MINUTE，否则取当前时刻。
    suffix 为空时覆盖 hourly_chart.bmp，否则写 hourly_chart_<suffix>.bmp
    """
    if not DATA_FILE.exists():
        print("  [图表] 无 CSV 数据，跳过")
        return None

    try:
        with open(DATA_FILE, 'r', encoding='utf-8-sig') as f:
            rows = list(csv.DictReader(f))
    except Exception:
        print("  [图表] CSV 读取失败，跳过")
        return None

    if len(rows) < 2:
        print("  [图表] 数据不足（需 ≥2 条记录），跳过")
        return None

    records = []
    for r in rows:
        try:
            t = datetime.strptime(r['timestamp'], '%Y-%m-%d %H:%M:%S')
            k = float(r['kwh'])
            records.append((t, k))
        except (ValueError, KeyError):
            continue
    if len(records) < 2:
        print("  [图表] 有效记录不足，跳过")
        return None
    records.sort(key=lambda x: x[0])

    now = datetime.now()
    if suffix:
        anchor = now.replace(hour=DAILY_SUMMARY_HOUR, minute=DAILY_SUMMARY_MINUTE,
                             second=0, microsecond=0)
    else:
        anchor = now
    end_time = anchor
    start_time = end_time - timedelta(hours=24)

    hourly = {}
    for t, k in records:
        if t < start_time or t > end_time:
            continue
        hour_key = t.replace(minute=0, second=0, microsecond=0)
        if hour_key not in hourly:
            hourly[hour_key] = {"first": (t, k), "last": (t, k)}
        else:
            if t < hourly[hour_key]["first"][0]:
                hourly[hour_key]["first"] = (t, k)
            if t > hourly[hour_key]["last"][0]:
                hourly[hour_key]["last"] = (t, k)

    hour_consumption = {}
    for hk, hd in hourly.items():
        first_kwh = hd["first"][1]
        last_kwh = hd["last"][1]
        # 同一小时内电量反升（中途充值）时无法算出消耗，该小时按缺失处理
        if first_kwh is not None and last_kwh is not None and first_kwh >= last_kwh:
            hour_consumption[hk] = round(first_kwh - last_kwh, 4)

    if not hour_consumption:
        print("  [图表] 时间范围内无有效小时数据，跳过")
        return None

    hours_list = []
    h = start_time.replace(minute=0, second=0, microsecond=0)
    while h < end_time:
        hours_list.append(h)
        h += timedelta(hours=1)

    values = [hour_consumption.get(h, 0) for h in hours_list]
    max_val = max(values) if any(values) else 1

    W, H = 800, 400
    TOP, BOT, LFT, RGT = 32, 52, 60, 16
    cw, ch_h = W - LFT - RGT, H - TOP - BOT
    bar_w = max(4, int(cw / len(hours_list) * 0.7))
    gap = cw / len(hours_list)
    total_kwh = sum(v for v in values if v)

    FONT = {
        '0': (0x0E,0x11,0x13,0x15,0x19,0x11,0x0E), '1': (0x04,0x0C,0x04,0x04,0x04,0x04,0x0E),
        '2': (0x0E,0x11,0x01,0x02,0x04,0x08,0x1F), '3': (0x0E,0x11,0x01,0x06,0x01,0x11,0x0E),
        '4': (0x02,0x06,0x0A,0x12,0x1F,0x02,0x02), '5': (0x1F,0x10,0x1E,0x01,0x01,0x11,0x0E),
        '6': (0x0E,0x10,0x1E,0x11,0x11,0x11,0x0E), '7': (0x1F,0x01,0x02,0x04,0x08,0x08,0x08),
        '8': (0x0E,0x11,0x11,0x0E,0x11,0x11,0x0E), '9': (0x0E,0x11,0x11,0x0F,0x01,0x01,0x0E),
        ':': (0x00,0x04,0x00,0x00,0x04,0x00,0x00), '/': (0x01,0x02,0x02,0x04,0x04,0x08,0x10),
        '.': (0x00,0x00,0x00,0x00,0x00,0x00,0x0C), ' ': (0x00,0x00,0x00,0x00,0x00,0x00,0x00),
        'k': (0x10,0x12,0x14,0x18,0x14,0x12,0x10), 'W': (0x11,0x11,0x11,0x15,0x15,0x1B,0x11),
        'h': (0x10,0x10,0x16,0x19,0x11,0x11,0x11), 't': (0x08,0x1C,0x08,0x08,0x08,0x09,0x06),
        'o': (0x00,0x00,0x0E,0x11,0x11,0x11,0x0E), 'a': (0x00,0x00,0x0C,0x02,0x0E,0x12,0x0E),
        'l': (0x0C,0x04,0x04,0x04,0x04,0x04,0x0E), '|': (0x04,0x04,0x04,0x04,0x04,0x04,0x04),
        'w': (0x00,0x00,0x11,0x11,0x15,0x15,0x0A), 's': (0x00,0x00,0x0F,0x10,0x0E,0x01,0x1E),
        '=': (0x00,0x00,0x1F,0x00,0x1F,0x00,0x00), '-': (0x00,0x00,0x00,0x1F,0x00,0x00,0x00),
        '>': (0x08,0x04,0x02,0x01,0x02,0x04,0x08), '(': (0x04,0x08,0x08,0x08,0x08,0x08,0x04),
        ')': (0x08,0x04,0x04,0x04,0x04,0x04,0x08), '~': (0x00,0x04,0x08,0x00,0x00,0x00,0x00),
        'A': (0x0E,0x11,0x11,0x1F,0x11,0x11,0x11), 'B': (0x1E,0x11,0x11,0x1E,0x11,0x11,0x1E),
        'C': (0x0E,0x11,0x10,0x10,0x10,0x11,0x0E), 'D': (0x1E,0x11,0x11,0x11,0x11,0x11,0x1E),
        'E': (0x1F,0x10,0x10,0x1E,0x10,0x10,0x1F), 'F': (0x1F,0x10,0x10,0x1E,0x10,0x10,0x10),
        'G': (0x0E,0x11,0x10,0x17,0x11,0x11,0x0E), 'H': (0x11,0x11,0x11,0x1F,0x11,0x11,0x11),
        'I': (0x0E,0x04,0x04,0x04,0x04,0x04,0x0E), 'J': (0x07,0x02,0x02,0x02,0x02,0x12,0x0C),
        'K': (0x11,0x12,0x14,0x18,0x14,0x12,0x11), 'L': (0x10,0x10,0x10,0x10,0x10,0x10,0x1F),
        'M': (0x11,0x1B,0x15,0x15,0x11,0x11,0x11), 'N': (0x11,0x19,0x15,0x13,0x11,0x11,0x11),
        'O': (0x0E,0x11,0x11,0x11,0x11,0x11,0x0E), 'P': (0x1E,0x11,0x11,0x1E,0x10,0x10,0x10),
        'Q': (0x0E,0x11,0x11,0x11,0x15,0x12,0x0D), 'R': (0x1E,0x11,0x11,0x1E,0x14,0x12,0x11),
        'S': (0x0F,0x10,0x10,0x0E,0x01,0x01,0x1E), 'T': (0x1F,0x04,0x04,0x04,0x04,0x04,0x04),
        'U': (0x11,0x11,0x11,0x11,0x11,0x11,0x0E), 'V': (0x11,0x11,0x11,0x11,0x11,0x0A,0x04),
        'X': (0x11,0x11,0x0A,0x04,0x0A,0x11,0x11), 'Y': (0x11,0x11,0x0A,0x04,0x04,0x04,0x04),
        'Z': (0x1F,0x01,0x02,0x04,0x08,0x10,0x1F),
    }
    FW, FH, FS = 5, 7, 1

    stride = (W * 3 + 3) // 4 * 4
    px = bytearray(stride * H)

    def _ps(x, y, r, g, b):
        if 0 <= x < W and 0 <= y < H:
            off = (H - 1 - y) * stride + x * 3
            px[off], px[off+1], px[off+2] = b, g, r

    def _rect(x, y, w, h, r, g, b):
        for dy in range(h):
            for dx in range(w):
                _ps(x+dx, y+dy, r, g, b)

    def _ch(x, y, c, r, g, b):
        gx = FONT.get(c)
        if gx is None: return
        for row in range(FH):
            bits = gx[row]
            for col in range(FW):
                if bits & (1 << (FW - 1 - col)):
                    _ps(x+col, y+row, r, g, b)

    def _txt(x, y, s, r, g, b):
        for i, c in enumerate(s):
            _ch(x + i*(FW+FS), y, c, r, g, b)

    _rect(0, 0, W, H, 0x2A, 0x17, 0x0F)

    # 标题行1: 时间范围 + 总消耗
    title = (f'{start_time.strftime("%m/%d %H:%M")} ~> {end_time.strftime("%m/%d %H:%M")}'
             f' | total={total_kwh:.2f}kWh')
    _txt(8, 4, title, 0xB8, 0xA3, 0x94)

    # 标题行2: 当前剩余电量 + 状态 + 更新时间
    try:
        latest = json.loads(LATEST_FILE.read_text(encoding='utf-8'))
        kwh_now = latest.get('kwh', '?')
        st_en = {'在线': 'ONLINE', '离线': 'OFFLINE'}.get(latest.get('status', ''), 'UNKNOWN')
        ts_now = str(latest.get('timestamp', '?'))[:16]
        _txt(8, 14, f'kwh={kwh_now} | {st_en} | ts={ts_now}', 0xF0, 0xE8, 0xE2)
    except Exception:
        pass

    # Y 轴单位
    _txt(2, TOP - 8, 'kWh', 0x69, 0x55, 0x47)

    y_steps = 4
    for i in range(y_steps + 1):
        yv = max_val * i / y_steps
        ys = TOP + ch_h - int(ch_h * i / y_steps)
        for px_x in range(LFT, W - RGT):
            _ps(px_x, ys, 0x3B, 0x29, 0x1E)
        lb = f'{yv:.2f}'
        _txt(LFT - len(lb)*(FW+FS) - 3, ys-3, lb, 0x69, 0x55, 0x47)

    for idx, hk in enumerate(hours_list):
        val = hour_consumption.get(hk, 0)
        hd = hourly.get(hk)
        if hd:
            span_min = int((hd["last"][0] - hd["first"][0]).total_seconds() // 60)
        else:
            span_min = 0
        bh = int(val / max_val * ch_h) if max_val > 0 else 0
        bx = int(LFT + idx * gap + (gap - bar_w) / 2)
        by = TOP + ch_h - bh

        # 柱色按该小时耗电量分档：>0.5kWh 蓝，>0.2kWh 青，其余橙
        if val > 0.5:   cr, cg, cb = 0x44, 0x44, 0xEF
        elif val > 0.2: cr, cg, cb = 0x08, 0xB3, 0xEA
        else:           cr, cg, cb = 0xF6, 0x82, 0x3B

        if bh > 0:
            _rect(bx, by, bar_w, bh, cr, cg, cb)

        if val > 0:
            vl = f'{val:.3f}'
            _txt(bx + bar_w//2 - len(vl)*3, by - 10, vl, 0xF0, 0xE8, 0xE2)

        # 柱底数字 = 该小时数据点实际覆盖的分钟数：0=蓝(单点采样)，<55=青(采集中断)，≥55=灰(完整)
        cnt = str(span_min)
        if span_min == 0:   mr, mg, mb = 0x44, 0x44, 0xEF
        elif span_min < 55: mr, mg, mb = 0x08, 0xB3, 0xEA
        else:               mr, mg, mb = 0x69, 0x55, 0x47
        _txt(bx + bar_w//2 - len(cnt)*3, TOP + ch_h + 4, cnt, mr, mg, mb)

        if idx % 2 == 0:
            hl = hk.strftime("%d/%H")
            _txt(bx + bar_w//2 - len(hl)*3, H - BOT + 14, hl, 0x69, 0x55, 0x47)

    bmp_size = 14 + 40 + len(px)
    bmp_hdr = struct.pack('<2sIHHI', b'BM', bmp_size, 0, 0, 54)
    dib_hdr = struct.pack('<IiiHHIIiiII', 40, W, H, 1, 24, 0, len(px), 2835, 2835, 0, 0)

    CHART_DIR.mkdir(parents=True, exist_ok=True)
    fname = f"hourly_chart_{suffix}.bmp" if suffix else "hourly_chart.bmp"
    chart_path = CHART_DIR / fname
    with open(chart_path, 'wb') as f:
        f.write(bmp_hdr + dib_hdr + px)

    n_valid = len([v for v in values if v > 0])
    print(f"  [图表] 已生成: {chart_path} ({len(hours_list)}小时, {n_valid}个有效, {bmp_size}B)")
    return chart_path


# ====================== 主流程 ======================
def run_once():
    """完整的单轮采集：抓取 → 存盘 → 告警 → 必要时出图。"""
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    t_total_start = time.time()

    # ── 0. 保留期清理（每天一次，CSV/图表只留最近 RETENTION_DAYS 天） ──
    try:
        prune_storage()
    except Exception as e:
        print(f"[{ts}] [WARN] 保留期清理异常: {type(e).__name__}: {e}")

    # ── 1. 抓取数据 ──
    t0 = time.time()
    try:
        kwh, status = fetch_meter_data()
        t_fetch = time.time() - t0
    except Exception as e:
        t_fetch = time.time() - t0
        print(f"[{ts}] [FAIL] 抓取失败 (耗时 {t_fetch:.2f}s): {type(e).__name__}: {e}")
        sys.exit(1)

    # ── 2. 保存数据 ──
    t1 = time.time()
    try:
        save_data(kwh, status)
        t_save = time.time() - t1
    except Exception as e:
        t_save = time.time() - t1
        print(f"[{ts}] [FAIL] 保存失败 (耗时 {t_save:.2f}s): {type(e).__name__}: {e}")
        sys.exit(2)

    # ── 3. 告警检查 ──
    t2 = time.time()
    alert_types = []
    try:
        alerts = check_alerts(kwh, status)
        if alerts:
            print(f"═══════════════════════════════════════")
            print(f"[{ts}] 采集 | kwh={kwh}  status={status}"
                  f"  fetch={t_fetch:.2f}s  save={t_save:.2f}s")
            print(f"── 告警详情 ──")
            for atype, title, content, nid, ongoing in alerts:
                if nid:
                    print(f"  [{atype}] {title}")
                    if notify(title, content, nid, ongoing):
                        print(f"    通知 [OK]")
                        alert_types.append(atype)
                    else:
                        print(f"    通知 [SKIP] Toast 发送失败（见 data/alerts.log）")
                else:
                    print(f"  [{atype}] {title}")
                    alert_types.append(atype)
            t_alert = time.time() - t2
            t_total = time.time() - t_total_start
            print(f"  耗时: alert={t_alert:.2f}s  total={t_total:.2f}s")
            print(f"═══════════════════════════════════════")
        else:
            t_alert = time.time() - t2
            t_total = time.time() - t_total_start
            print(f"[{ts}] kwh={kwh} | {status}"
                  f" | {t_fetch:.1f}s/{t_save:.1f}s/{t_alert:.1f}s/{t_total:.1f}s")

        if alert_types:
            update_alert_state(alert_types, status)
        else:
            state = load_alert_state()
            state['last_status'] = status
            state['last_check'] = datetime.now().isoformat()
            save_alert_state(state)

    except Exception as e:
        t_alert = time.time() - t2
        print(f"[{ts}] [WARN] 告警检查异常 (耗时 {t_alert:.2f}s): {e}")

    # ── 4. 生成图表 ──
    # 只在「整点那一分钟」生成一次；每分钟生成会反复全量读 CSV + 纯 Python 绘 800x400 BMP，
    # 连续的 CPU 尖峰没有意义（一小时内数据变化不足以看清）
    if datetime.now().minute == 0:
        try:
            generate_hourly_chart()
        except Exception as e:
            print(f"  [图表] 实时图表生成失败: {e}")

    # 每日汇总触发时额外存一份快照（快照同样受 RETENTION_DAYS 保留期约束）
    if any(at == 'daily_summary' for at in alert_types):
        daily_suffix = datetime.now().strftime('%Y%m%d_%H%M')
        try:
            generate_hourly_chart(suffix=daily_suffix)
        except Exception as e:
            print(f"  [图表] 每日快照生成失败: {e}")


def _sleep_to_next_minute():
    """睡到下一个整分，保证整点（minute == 0）那一分钟一定会被命中。"""
    now = datetime.now()
    nxt = (now + timedelta(minutes=1)).replace(second=0, microsecond=0)
    time.sleep(max(1.0, (nxt - now).total_seconds()))


def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="电表电量监控（Windows PC 版）：抓电量、存 CSV、Toast 告警、画 BMP 图表",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--loop", action="store_true",
                   help="常驻运行，按整分钟对齐每分钟采集一次")
    p.add_argument("--chart", action="store_true",
                   help="只生成当前 24 小时图表后退出（调试用）")
    p.add_argument("--test-notify", action="store_true",
                   help="发一条测试 Toast 后退出，验证通知链路")
    p.add_argument("--prune", action="store_true",
                   help=f"立刻按保留期（{RETENTION_DAYS} 天）清理 CSV 与图表后退出")
    p.add_argument("--no-log", action="store_true", help="不写 data/monitor.log")
    return p.parse_args(argv)


def main(argv=None):
    setup_console()
    args = parse_args(argv)
    if not args.no_log:
        setup_log()

    if args.test_notify:
        ok = notify("电表监控 - 测试通知",
                    f"表号: {METER_NO}\n看到这条通知说明通知链路正常。", NID_TEST)
        print(f"[测试] Toast {'发送成功' if ok else '发送失败，详见 data/alerts.log'}")
        return 0 if ok else 3

    if args.chart:
        path = generate_hourly_chart()
        print(f"[图表] {path if path else '未生成（数据不足）'}")
        return 0 if path else 4

    if args.prune:
        prune_storage(force=True)
        print(f"[清理] 完成（保留 {RETENTION_DAYS} 天）")
        return 0

    if not args.loop:
        run_once()
        return 0

    print(f"[loop] 启动常驻模式 PID={os.getpid()}，每分钟整分采集一次，Ctrl+C 退出")
    while True:
        try:
            run_once()
        except SystemExit as e:      # run_once 里抓取/存盘失败会 sys.exit
            print(f"[loop] 本轮提前结束 (exit={e.code})，下一分钟继续")
        except Exception as e:
            print(f"[loop] 未捕获异常: {type(e).__name__}: {e}")
        try:
            _sleep_to_next_minute()
        except KeyboardInterrupt:
            print("\n[loop] 收到 Ctrl+C，退出")
            return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n[退出] 已中断")
        sys.exit(130)
