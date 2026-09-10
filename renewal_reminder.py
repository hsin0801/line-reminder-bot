import os
import io
import json
import requests
from datetime import datetime, date, timedelta, timezone

LINE_TOKEN        = os.environ.get("LINE_TOKEN")
TARGET_GROUP_ID   = os.environ.get("REMINDER_GROUP_ID")
RENEWAL_FILE_ID   = os.environ.get("RENEWAL_SHEET_ID", "1-6Wmly1lKSLVOEUspwcV9TPsghdjyMC_")
REMIND_START_HOUR = 9
REMIND_END_HOUR   = 20
MEMBER_USER_IDS   = json.loads(os.environ.get("MEMBER_USER_IDS", "{}"))
TW_TZ             = timezone(timedelta(hours=8))

# ── 狀態改存 Drive ───────────────────────────────────
# 原本 STATE_FILE = "reminder_state.json" 存在 Render 本機硬碟。
# Render 免費方案閒置 15 分鐘就休眠，醒來後本機檔案全被清空，
# load_state() 會拿到空 dict，於是每個人的 reminded_today 都變回 False，
# 下一次 hourly cron 又整輪重新提醒一次——這就是 2026-09 額度被燒光的原因。
# 改成跟 renewal_progress_push.py 一樣存 Drive，休眠與重新部署都不會掉。
STATE_FOLDER_ID   = os.environ.get(
    "RENEWAL_PROGRESS_FOLDER_ID", "1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf"
)
STATE_FILENAME    = "reminder_state.json"

_holiday_cache = {}

def get_holidays(year: int) -> set:
    if year in _holiday_cache:
        return _holiday_cache[year]
    holidays = set()
    try:
        url = f"https://data.ntpc.gov.tw/api/datasets/308DCD75-6119-4125-8843-2057C0E43ED5/json?$top=500&$filter=year%20eq%20{year}"
        resp = requests.get(url, timeout=10, verify=False)
        if resp.status_code == 200:
            try:
                data = resp.json()
                for item in data:
                    d = item.get("date", "")
                    if len(d) == 8:
                        holidays.add(date(int(d[:4]), int(d[4:6]), int(d[6:])))
            except:
                pass
    except Exception as e:
        print(f"[WARNING] 假日 API 失敗: {e}")
    _holiday_cache[year] = holidays
    return holidays

def is_off_day(d: date) -> bool:
    if d.weekday() == 6:
        return True
    return d in get_holidays(d.year)

def next_work_day(d: date) -> date:
    next_d = d
    while is_off_day(next_d):
        next_d += timedelta(days=1)
    return next_d

def get_trigger_date(month_day: int, year: int, month: int) -> date:
    try:
        original = date(year, month, month_day)
    except ValueError:
        return None
    return next_work_day(original)

def get_days_left_in_month(d: date) -> int:
    if d.month == 12:
        last_day = date(d.year + 1, 1, 1) - timedelta(days=1)
    else:
        last_day = date(d.year, d.month + 1, 1) - timedelta(days=1)
    return (last_day - d).days + 1

# ── 讀取續保資料（Drive API 下載 xlsx）──────────────
def get_renewal_data() -> dict:
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        import openpyxl

        creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
        if not creds_json:
            print("[ERROR] 缺少 GOOGLE_CREDENTIALS_JSON")
            return {}

        creds_dict = json.loads(creds_json)
        scopes = ["https://www.googleapis.com/auth/drive.readonly"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        service = build("drive", "v3", credentials=creds)

        request = service.files().get_media(fileId=RENEWAL_FILE_ID)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        buf.seek(0)
        wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)

        today = date.today()
        ws = None
        for name in wb.sheetnames:
            if f"{today.month:02d}月" in name or f"{today.month}月" in name:
                ws = wb[name]
                break
        if not ws:
            ws = wb.active

        skip_names = {"劉珈微", "劉宗鑫", "歸仁一課", "歸仁二課",
                      "合計", "營業員", "None", ""}

        def safe_int(val):
            try:
                if val is None or str(val).strip() in ("", "-", "None"):
                    return 0
                return int(float(str(val).replace(",", "").strip()))
            except:
                return 0

        result = {}
        headers = []
        header_found = False

        for row in ws.iter_rows(values_only=True):
            row_vals = [str(v).strip() if v is not None else "" for v in row]
            if not header_found:
                if "營業員" in row_vals and "母數" in row_vals:
                    headers = row_vals
                    header_found = True
                continue

            name = row_vals[0] if row_vals else ""
            if not name or name in skip_names:
                continue

            def col(keyword):
                for i, h in enumerate(headers):
                    if keyword in h:
                        return safe_int(row_vals[i]) if i < len(row_vals) else 0
                return 0

            mu = col("母數")
            if mu == 0:
                continue

            result[name] = {
                "母數":     mu,
                "預估":     col("預估"),
                "已收":     col("已收"),
                "續保率":   col("已收") / mu if mu > 0 else 0,
                "首年母數": col("首年續保母") or col("首年母"),
                "首年預估": col("首年續保預") or col("首年預"),
                "首年已收": col("首年續保已") or col("首年已"),
                "車體母數": col("車體母"),
                "車體預估": col("車體預"),
                "車體已收": col("車體已"),
            }

        wb.close()
        print(f"[OK] 讀取續保資料成功，共 {len(result)} 人")
        return result

    except Exception as e:
        print(f"[ERROR] 讀取續保資料失敗: {e}")
        return {}

# ── 提醒條件判斷 ────────────────────────────────────
def should_remind(person_data: dict, today: date):
    day    = today.day
    mu     = person_data["母數"]
    est    = person_data["預估"]
    col    = person_data["已收"]
    rate   = person_data["續保率"]

    if mu == 0:
        return False, ""

    y, m = today.year, today.month
    t1 = get_trigger_date(1,  y, m)
    t2 = get_trigger_date(7,  y, m)
    t3 = get_trigger_date(15, y, m)
    t4 = get_trigger_date(20, y, m)

    if today == t1 and rate < 0.20:
        gap = max(0, round(mu * 0.20) - col)
        return True, f"月初進度僅 {rate*100:.1f}%（{col}/{mu}台），距門檻20%還差 {gap} 台"
    if today == t2 and rate < 0.50:
        gap = max(0, round(mu * 0.50) - col)
        return True, f"進度僅 {rate*100:.1f}%（{col}/{mu}台），距門檻50%還差 {gap} 台"
    if today == t3 and est > 0 and col < est:
        return True, f"進度 {rate*100:.1f}%（{col}/{mu}台），落後預估 {est-col} 台"
    if today == t4 and est > 0 and col < est:
        return True, f"⚠️ 第二次警示：進度 {rate*100:.1f}%（{col}/{mu}台），落後預估 {est-col} 台"
    if day >= 21 and est > 0 and col < est:
        days_left = get_days_left_in_month(today)
        return True, f"距月底剩 {days_left} 天，進度 {rate*100:.1f}%（{col}/{mu}台），還差 {est-col} 台"

    return False, ""

# ── LINE 發送 ────────────────────────────────────────
def _utf16_index(text: str, sub: str) -> int:
    """
    LINE 的 mention index 是以 UTF-16 code unit 計算的，不是 Python 的字元數。
    像 📊 這種 BMP 以外的字元，Python 算 1 個字、UTF-16 算 2 個單位。
    原本直接用 text.index() 會整串偏移，標註會圈到錯的字。
    """
    i = text.index(sub)
    return len(text[:i].encode("utf-16-le")) // 2


def build_batch_message(items: list, is_followup: bool = False) -> dict:
    """
    把「這一輪所有需要提醒的人」組成【一則】訊息，而不是一人一則。

    為什麼：LINE 是以發送對象人數計費，推到 9 人群組一次算 9 則。
    原本一人一則，5 個人要提醒就是 5 次推播 = 45 則；
    合併成一則只花 9 則，省 5 倍。人多的時候差更多。

    items: [(name, reason), ...]
    """
    month = date.today().month

    if is_followup:
        header = f"⚠️ {month}月續保進度：以下同仁尚未回覆"
        lines  = [f"@{name}" for name, _ in items]
        footer = "請盡快說明追蹤計畫 🙏"
    else:
        header = f"📊 {month}月續保進度提醒"
        lines  = [f"@{name} {reason}" for name, reason in items]
        footer = "請今天內回覆追蹤計畫 🙏"

    text = header + "\n\n" + "\n".join(lines) + "\n\n" + footer
    msg  = {"type": "text", "text": text}

    mentionees = []
    for name, _ in items:
        user_id = MEMBER_USER_IDS.get(name)
        if not user_id:
            continue
        at = f"@{name}"
        try:
            mentionees.append({
                "index":  _utf16_index(text, at),
                "length": len(at.encode("utf-16-le")) // 2,
                "type":   "user",
                "userId": user_id,
            })
        except ValueError:
            continue

    if mentionees:
        msg["mention"] = {"mentionees": mentionees}
    return msg


def push_batch(items: list, is_followup: bool = False) -> bool:
    """
    送出合併後的提醒。送出前先問額度守門員夠不夠，不夠就不送。
    回傳是否真的送出去了。
    """
    if not items:
        return False

    import line_quota

    label = "續保追蹤提醒" if is_followup else "續保進度提醒"
    ok, info = line_quota.check(TARGET_GROUP_ID, pushes=1, priority="normal")
    if not ok:
        print(
            f"[SKIP] {label} 因額度不足未發送："
            f"本月已用 {info['used']}/{info['limit']} 則，"
            f"一般推播還剩 {info['remaining_pushes']} 次"
        )
        return False

    msg = build_batch_message(items, is_followup=is_followup)
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    try:
        resp = requests.post(url, headers=headers,
                             json={"to": TARGET_GROUP_ID, "messages": [msg]},
                             timeout=10)
    except Exception as e:
        print(f"[ERROR] {label} 發送例外: {e}")
        return False

    names = "、".join(n for n, _ in items)
    print(f"[LINE] {label} @{names} 發送: {resp.status_code}")

    if resp.status_code != 200:
        # 429 = 月額度用完。這種失敗一定要看得見，不然又是整月靜默。
        print(f"[ERROR] {label} 發送失敗 HTTP {resp.status_code}: {resp.text[:200]}")
        return False

    line_quota.commit(info, pushes_sent=1, label=label)
    return True


# ── 狀態管理（存 Drive，不存本機）────────────────────
def load_state() -> dict:
    from drive_json_store import load_json_from_drive
    try:
        return load_json_from_drive(STATE_FOLDER_ID, STATE_FILENAME) or {}
    except Exception as e:
        # 讀不到就回空 dict —— 行為跟以前一樣，但至少會留下紀錄
        print(f"[WARN] 讀取 {STATE_FILENAME} 失敗，本輪視為無狀態: {e}")
        return {}

def save_state(state: dict):
    # 註：服務帳戶沒有自己的儲存配額，無法在 My Drive 資料夾「新建」檔案，
    # 只能更新既有檔案。所以 reminder_state.json 必須由 Hsin 的帳號預先建立一次，
    # 之後 Bot 才寫得進去。不要刪掉它。
    from drive_json_store import save_json_to_drive
    try:
        save_json_to_drive(STATE_FOLDER_ID, STATE_FILENAME, state)
    except Exception as e:
        # 寫不進去會讓防重複失效（下輪又整輪重發），一定要吼出來
        print(f"[ERROR] 寫入 {STATE_FILENAME} 失敗，防重複會失效: {e}")

def reset_daily_state(state: dict, today_str: str) -> dict:
    for name in state:
        if state[name].get("last_reset") != today_str:
            state[name]["replied_today"]  = False
            state[name]["reminded_today"] = False
            state[name]["followup_count"] = 0        # ← 加這行
            state[name]["last_reset"]     = today_str
    return state

# ── 主入口 ───────────────────────────────────────────
def run_reminder():
    now          = datetime.now(TW_TZ)
    today        = now.date()
    today_str    = today.isoformat()
    current_hour = now.hour

    if current_hour < REMIND_START_HOUR or current_hour >= REMIND_END_HOUR:
        print(f"[SKIP] 目前台灣時間 {current_hour}:00，不在提醒時間內（{REMIND_START_HOUR}-{REMIND_END_HOUR}）")
        return

    if today.weekday() == 6:
        print("[SKIP] 今天是週日，跳過")
        return

    if today in get_holidays(today.year):
        print("[SKIP] 今天是國定假日，跳過")
        return

    renewal_data = get_renewal_data()
    if not renewal_data:
        print("[SKIP] 無法取得續保資料")
        return

    state = load_state()
    state = reset_daily_state(state, today_str)

    # 先「收集」這一輪誰該被提醒，全部跑完再合併成一則送出去。
    # 原本是在迴圈裡一人送一次，5 個人就是 5 次推播 × 群組人數；
    # 改成收集後合併，同樣 5 個人只花 1 次推播。
    first_time = []   # [(name, reason), ...] 今天第一次提醒
    followups  = []   # [(name, reason), ...] 已提醒過但還沒回覆

    for name, data in renewal_data.items():
        if name not in state:
            state[name] = {
                "replied_today":  False,
                "reminded_today": False,
                "last_reset":     today_str,
                "last_remind_time": None
            }

        person = state[name]
        if person.get("replied_today"):
            continue

        need_remind, reason = should_remind(data, today)
        if not need_remind:
            continue

        last_remind = person.get("last_remind_time")

        if not person.get("reminded_today"):
            if REMIND_START_HOUR <= current_hour <= 17:
                first_time.append((name, reason))
        else:
            if last_remind:
                last_dt     = datetime.fromisoformat(last_remind)
                hours_since = (now - last_dt).total_seconds() / 3600
                if hours_since >= 2 and person.get("followup_count", 0) < 2:
                    followups.append((name, reason))

    # 送出（各自最多一次推播），成功才更新狀態——
    # 若因額度不足沒送出，狀態不動，額度恢復後還會再試。
    if first_time and push_batch(first_time, is_followup=False):
        for name, _ in first_time:
            state[name]["reminded_today"]   = True
            state[name]["last_remind_time"] = now.isoformat()
        print(f"[REMIND] 首次提醒 {len(first_time)} 人：" + "、".join(n for n, _ in first_time))

    if followups and push_batch(followups, is_followup=True):
        for name, _ in followups:
            state[name]["last_remind_time"] = now.isoformat()
            state[name]["followup_count"]   = state[name].get("followup_count", 0) + 1
        print(f"[FOLLOWUP] 追蹤提醒 {len(followups)} 人：" + "、".join(n for n, _ in followups))

    save_state(state)
    print(f"[DONE] {now.strftime('%Y-%m-%d %H:%M')} 台灣時間，提醒任務完成")

def mark_replied(user_display_name: str):
    state     = load_state()
    today_str = date.today().isoformat()
    for name in state:
        if name in user_display_name or user_display_name in name:
            state[name]["replied_today"] = True
            save_state(state)
            print(f"[REPLIED] {name} 已回覆")
            return True
    return False
