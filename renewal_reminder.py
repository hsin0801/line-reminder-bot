import os
import json
import requests
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────
#  設定區
# ─────────────────────────────────────────
LINE_TOKEN        = os.environ.get("LINE_TOKEN")
TARGET_GROUP_ID   = os.environ.get("REMINDER_GROUP_ID")   # 單一提醒群組
SPREADSHEET_ID    = os.environ.get("RENEWAL_SHEET_ID")     # Google Sheets ID
HOLIDAY_API_YEAR  = date.today().year
REMIND_START_HOUR = 9    # 早上9點開始
REMIND_END_HOUR   = 20   # 晚上8點後停止

# 追蹤狀態檔案（存在本地或 Render Persistent Disk）
STATE_FILE = "reminder_state.json"

# 業務員對應的 LINE userId（需要事先取得，才能 @mention）
# 格式: { "顯示名稱": "LINE_USER_ID" }
MEMBER_USER_IDS = json.loads(os.environ.get("MEMBER_USER_IDS", "{}"))
# 範例環境變數值:
# {"林定緯":"Uxxxx","林適緯":"Uxxxx","陳建道":"Uxxxx",
#  "陳星佑":"Uxxxx","張姉瑀":"Uxxxx","歐陽文智":"Uxxxx","蔡明憬":"Uxxxx"}

# ─────────────────────────────────────────
#  假日判斷
# ─────────────────────────────────────────
_holiday_cache = {}

def get_holidays(year: int) -> set:
    """從政府開放資料 API 取得國定假日清單"""
    if year in _holiday_cache:
        return _holiday_cache[year]
    try:
        url = f"https://data.ntpc.gov.tw/api/datasets/308DCD75-6119-4125-8843-2057C0E43ED5/json?$top=500&$filter=year%20eq%20{year}"
        resp = requests.get(url, timeout=10, verify=False)
        holidays = set()
        if resp.status_code == 200:
            for item in resp.json():
                d = item.get("date", "")          # 格式 "20260101"
                if len(d) == 8:
                    holidays.add(date(int(d[:4]), int(d[4:6]), int(d[6:])))
        _holiday_cache[year] = holidays
        return holidays
    except Exception as e:
        print(f"[WARNING] 假日 API 失敗: {e}")
        return set()

def is_off_day(d: date) -> bool:
    """判斷是否為休息日（週日 或 國定假日）"""
    if d.weekday() == 6:   # 週日
        return True
    holidays = get_holidays(d.year)
    return d in holidays

def next_work_day(d: date) -> date:
    """取得下一個工作日（跳過週日與國定假日）"""
    next_d = d
    while is_off_day(next_d):
        next_d += timedelta(days=1)
    return next_d

def get_trigger_date(month_day: int, year: int, month: int) -> date:
    """
    取得觸發日（原定日期若為休息日則順延到下一工作日）
    month_day: 原定幾號
    """
    try:
        original = date(year, month, month_day)
    except ValueError:
        return None
    return next_work_day(original)

# ─────────────────────────────────────────
#  Google Sheets 讀取
# ─────────────────────────────────────────
def get_renewal_data() -> dict:
    """
    從 Google Sheets 讀取當月續保進度
    回傳格式:
    {
      "林定緯": {
        "母數": 29, "已收": 16, "預估": 27,
        "續保率": 0.552,
        "首年母數": 5, "首年已收": 2, "首年預估": 4,
        "車體母數": 5, "車體已收": 1, "車體預估": 3
      },
      ...
    }
    """
    try:
        creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
        if not creds_json:
            print("[ERROR] 缺少 GOOGLE_CREDENTIALS_JSON 環境變數")
            return {}

        creds_dict = json.loads(creds_json)
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)

        sh = gc.open_by_key(SPREADSHEET_ID)
        # 試圖找到當月的工作表（例如 "115.06月續保"）
        today = date.today()
        roc_year = today.year - 1911
        sheet_name = f"{roc_year}.{today.month:02d}月續保"

        try:
            ws = sh.worksheet(sheet_name)
        except Exception:
            # fallback: 用第一個工作表
            ws = sh.get_worksheet(0)

        rows = ws.get_all_values()

        # 解析表頭找欄位位置
        header = []
        data_start = 0
        for i, row in enumerate(rows):
            if "營業員" in row and "母數" in row:
                header = row
                data_start = i + 1
                break

        if not header:
            print("[WARNING] 找不到表頭")
            return {}

        def col(name):
            try:
                return header.index(name)
            except ValueError:
                return None

        idx = {
            "name":     col("營業員"),
            "mu":       col("母數"),
            "estimate": col("預估"),
            "collected":col("已收"),
            "rate":     col("續保率"),
            "fy_mu":    header.index("首年續\n保母數")   if "首年續\n保母數"   in header else col("首年續 保母數"),
            "fy_est":   header.index("首年續\n保預估")   if "首年續\n保預估"   in header else col("首年續 保預估"),
            "fy_col":   header.index("首年續\n保已收")   if "首年續\n保已收"   in header else col("首年續 保已收"),
            "car_mu":   header.index("首年車\n體母數")   if "首年車\n體母數"   in header else col("首年車 體母數"),
            "car_est":  header.index("首年車\n體預估")   if "首年車\n體預估"   in header else col("首年車 體預估"),
            "car_col":  header.index("首年車體已收")     if "首年車體已收"     in header else None,
        }

        # 不顯示個人欄位的名單
        skip_names = {"劉珈微", "劉宗鑫", "歸仁一課", "歸仁二課", "合計", "營業員"}

        def safe_int(val):
            try:
                return int(str(val).replace(",", "").strip()) if val and str(val).strip() not in ("", "-") else 0
            except:
                return 0

        result = {}
        for row in rows[data_start:]:
            if not row or idx["name"] is None:
                continue
            name = str(row[idx["name"]]).strip()
            if not name or name in skip_names:
                continue

            mu  = safe_int(row[idx["mu"]])       if idx["mu"]        is not None else 0
            est = safe_int(row[idx["estimate"]]) if idx["estimate"]   is not None else 0
            col_val = safe_int(row[idx["collected"]]) if idx["collected"] is not None else 0

            if mu == 0:
                continue

            result[name] = {
                "母數":     mu,
                "預估":     est,
                "已收":     col_val,
                "續保率":   col_val / mu if mu > 0 else 0,
                "首年母數": safe_int(row[idx["fy_mu"]])  if idx["fy_mu"]  is not None else 0,
                "首年預估": safe_int(row[idx["fy_est"]]) if idx["fy_est"] is not None else 0,
                "首年已收": safe_int(row[idx["fy_col"]]) if idx["fy_col"] is not None else 0,
                "車體母數": safe_int(row[idx["car_mu"]])  if idx["car_mu"]  is not None else 0,
                "車體預估": safe_int(row[idx["car_est"]]) if idx["car_est"] is not None else 0,
                "車體已收": safe_int(row[idx["car_col"]]) if idx["car_col"] is not None else 0,
            }
        return result

    except Exception as e:
        print(f"[ERROR] 讀取 Sheets 失敗: {e}")
        return {}

# ─────────────────────────────────────────
#  提醒條件判斷
# ─────────────────────────────────────────
def should_remind(person_data: dict, today: date) -> tuple[bool, str]:
    """
    判斷是否需要提醒，回傳 (需要提醒, 原因說明)
    """
    day = today.day
    mu       = person_data["母數"]
    est      = person_data["預估"]
    col_val  = person_data["已收"]
    rate     = person_data["續保率"]

    if mu == 0:
        return False, ""

    # 計算各觸發日
    y, m = today.year, today.month
    t1 = get_trigger_date(1,  y, m)   # 月初：進度 < 20%
    t2 = get_trigger_date(7,  y, m)   # 7日：進度 < 50%
    t3 = get_trigger_date(15, y, m)   # 15日：低於預估
    t4 = get_trigger_date(20, y, m)   # 20日：低於預估（第二次）

    if today == t1 and rate < 0.20:
        gap = max(0, round(mu * 0.20) - col_val)
        return True, f"月初進度僅 {rate*100:.1f}%（{col_val}/{mu}台），距門檻20%還差 {gap} 台"

    if today == t2 and rate < 0.50:
        gap = max(0, round(mu * 0.50) - col_val)
        return True, f"進度僅 {rate*100:.1f}%（{col_val}/{mu}台），距門檻50%還差 {gap} 台"

    if today == t3 and est > 0 and col_val < est:
        gap = est - col_val
        return True, f"進度 {rate*100:.1f}%（{col_val}/{mu}台），落後預估 {gap} 台"

    if today == t4 and est > 0 and col_val < est:
        gap = est - col_val
        return True, f"⚠️ 第二次警示：進度 {rate*100:.1f}%（{col_val}/{mu}台），落後預估 {gap} 台"

    if day >= 21 and est > 0 and col_val < est:
        days_left = get_days_left_in_month(today)
        gap = est - col_val
        return True, f"距月底剩 {days_left} 天，進度 {rate*100:.1f}%（{col_val}/{mu}台），還差 {gap} 台"

    return False, ""

def get_days_left_in_month(d: date) -> int:
    """計算當月剩餘天數"""
    if d.month == 12:
        last_day = date(d.year + 1, 1, 1) - timedelta(days=1)
    else:
        last_day = date(d.year, d.month + 1, 1) - timedelta(days=1)
    return (last_day - d).days + 1

# ─────────────────────────────────────────
#  LINE 發送
# ─────────────────────────────────────────
def push_message(text: str):
    """推播訊息到指定群組"""
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    body = {
        "to": TARGET_GROUP_ID,
        "messages": [{"type": "text", "text": text}]
    }
    resp = requests.post(url, headers=headers, json=body)
    print(f"[LINE] 發送結果: {resp.status_code} {resp.text[:100]}")

def build_mention_message(name: str, reason: str, is_followup: bool = False) -> dict:
    """
    建立帶有 @mention 的訊息
    回傳 LINE message object
    """
    user_id = MEMBER_USER_IDS.get(name)
    today = date.today()
    month = today.month

    if user_id:
        mention = f"@{name}"
        if is_followup:
            text = f"{mention} 尚未收到你的回覆\n請盡快說明 {month}月續保進度狀況 🙏"
        else:
            text = f"📊 {month}月續保進度提醒\n{mention}\n{reason}\n\n請今天內回覆追蹤計畫 🙏"

        return {
            "type": "text",
            "text": text,
            "mention": {
                "mentionees": [{
                    "index": text.index(mention),
                    "length": len(mention),
                    "type": "user",
                    "userId": user_id
                }]
            }
        }
    else:
        # 沒有 userId，用純文字
        if is_followup:
            text = f"@{name} 尚未收到你的回覆\n請盡快說明 {month}月續保進度狀況 🙏"
        else:
            text = f"📊 {month}月續保提醒\n@{name}\n{reason}\n\n請今天內回覆追蹤計畫 🙏"
        return {"type": "text", "text": text}

def push_mention(name: str, reason: str, is_followup: bool = False):
    """推播 @mention 訊息"""
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    msg = build_mention_message(name, reason, is_followup)
    body = {"to": TARGET_GROUP_ID, "messages": [msg]}
    resp = requests.post(url, headers=headers, json=body)
    print(f"[LINE] @{name} 發送結果: {resp.status_code}")

# ─────────────────────────────────────────
#  狀態管理
# ─────────────────────────────────────────
def load_state() -> dict:
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_state(state: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def reset_daily_state(state: dict, today_str: str) -> dict:
    """每天重置當日回覆狀態"""
    for name in state:
        if state[name].get("last_reset") != today_str:
            state[name]["replied_today"] = False
            state[name]["reminded_today"] = False
            state[name]["last_reset"] = today_str
    return state

# ─────────────────────────────────────────
#  主入口：每小時由 Cron 觸發
# ─────────────────────────────────────────
def run_reminder():
    from datetime import timezone, timedelta
    TW_TZ = timezone(timedelta(hours=8))
    now = datetime.now(TW_TZ)
    today = now.date()
    today_str = today.isoformat()
    current_hour = now.hour

    # 1. 時間範圍檢查
    if current_hour < REMIND_START_HOUR or current_hour >= REMIND_END_HOUR:
        print(f"[SKIP] 目前 {current_hour}:00，不在提醒時間內（{REMIND_START_HOUR}-{REMIND_END_HOUR}）")
        return

    # 2. 週日跳過
    if today.weekday() == 6:
        print(f"[SKIP] 今天是週日，跳過")
        return

    # 3. 國定假日跳過
    if today in get_holidays(today.year):
        print(f"[SKIP] 今天是國定假日，跳過")
        return

    # 4. 讀取進度資料
    renewal_data = get_renewal_data()
    if not renewal_data:
        print("[SKIP] 無法取得續保資料")
        return

    # 5. 載入並重置狀態
    state = load_state()
    state = reset_daily_state(state, today_str)

    for name, data in renewal_data.items():
        if name not in state:
            state[name] = {
                "replied_today": False,
                "reminded_today": False,
                "last_reset": today_str,
                "last_remind_time": None
            }

        person = state[name]

        # 如果已回覆，跳過
        if person.get("replied_today"):
            continue

        # 判斷是否需要提醒
        need_remind, reason = should_remind(data, today)
        if not need_remind:
            continue

        last_remind = person.get("last_remind_time")

        if not person.get("reminded_today"):
            # 第一次提醒（今天還沒提醒過）：只在09:00發
            if current_hour == REMIND_START_HOUR:
                push_mention(name, reason, is_followup=False)
                person["reminded_today"] = True
                person["last_remind_time"] = now.isoformat()
                print(f"[REMIND] 首次提醒 @{name}")
        else:
            # 後續追蹤：每2小時一次
            if last_remind:
                last_dt = datetime.fromisoformat(last_remind)
                hours_since = (now - last_dt).total_seconds() / 3600
                if hours_since >= 2:
                    push_mention(name, reason, is_followup=True)
                    person["last_remind_time"] = now.isoformat()
                    print(f"[FOLLOWUP] 追蹤提醒 @{name}（距上次 {hours_since:.1f} 小時）")

    save_state(state)
    print(f"[DONE] {now.strftime('%Y-%m-%d %H:%M')} 提醒任務完成")

# ─────────────────────────────────────────
#  標記某人已回覆（由 webhook 呼叫）
# ─────────────────────────────────────────
def mark_replied(user_display_name: str):
    """
    當群組有人發言，找到對應的業務員名稱並標記已回覆
    user_display_name: LINE 顯示名稱
    """
    state = load_state()
    today_str = date.today().isoformat()

    # 找名字匹配（LINE 顯示名稱可能跟表格名稱不完全一樣，做模糊匹配）
    for name in state:
        if name in user_display_name or user_display_name in name:
            state[name]["replied_today"] = True
            save_state(state)
            print(f"[REPLIED] {name} 已回覆")
            return True
    return False
