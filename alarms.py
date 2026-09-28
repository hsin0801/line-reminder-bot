"""
群組訊息 → 私訊提醒
------------------------------------------------
只處理 ALARM_OWNER_ID（宗鑫）發的訊息：訊息裡有提到時間，就用 AI 解析出事件時間與內容，
存到 Drive 的 alarms.json，時間到前 LEAD_MINUTES 分鐘私訊提醒。
cron-job.org 每分鐘打 /check-alarms 送出到期的提醒。

alarms.json 必須由宗鑫的帳號預先建立（服務帳戶不能在 My Drive 新建檔案），不要刪掉。
"""

import os
import re
import json
import uuid
import threading
from datetime import datetime, timedelta, timezone

TW_TZ = timezone(timedelta(hours=8))
ALARM_OWNER_ID = os.environ.get("ALARM_OWNER_ID", "U272a3c6b1f3d10a3677769cb4f73fe1d")
LEAD_MINUTES = 10
FOLDER_ID = os.environ.get("RENEWAL_PROGRESS_FOLDER_ID", "1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf")
FILENAME = "alarms.json"
WEEKDAYS = "一二三四五六日"

# 沒提到時間的訊息不送 AI，省額度也避免誤判
TIME_HINT = re.compile(
    r"\d{1,2}\s*[:：]\s*\d{2}|\d{1,2}\s*[點点]|[上下]午|早上|中午|晚上|傍晚|今晚|今天|明天|後天|"
    r"週[一二三四五六日天]|星期|禮拜|\d{1,2}\s*/\s*\d{1,2}|\d{1,2}\s*月\s*\d{1,2}"
)

_lock = threading.Lock()


def _now():
    return datetime.now(TW_TZ)


def _fmt(dt):
    return f"{dt.month}/{dt.day}({WEEKDAYS[dt.weekday()]}) {dt:%H:%M}"


def _load():
    from drive_json_store import load_json_from_drive
    data = load_json_from_drive(FOLDER_ID, FILENAME) or {}
    return data if isinstance(data.get("alarms"), list) else {"alarms": []}


def _save(data):
    from drive_json_store import save_json_to_drive
    save_json_to_drive(FOLDER_ID, FILENAME, data)


def _pending(data):
    items = [a for a in data["alarms"] if not a.get("sent") and not a.get("cancelled")]
    return sorted(items, key=lambda a: a["event_at"])


PARSE_PROMPT = """現在時間是 {now}（台灣時間）。
判斷下面這則訊息是否在講「某個時間點要做的事」。是的話，解析出事件的日期時間和要做的事。
- 只說「今天」「下午5點」等，依現在時間推算日期；只說時間沒說日期，就當作今天（若該時間已過則當作明天）
- 沒有明確時間（只有日期或只說「月底」）就不算
- 只回傳 JSON，不要其他文字：
{{"is_task": true 或 false, "datetime": "YYYY-MM-DD HH:MM", "title": "要做的事（精簡，保留原意）"}}

訊息：{text}"""


def parse(text):
    """回傳 (event_dt, title) 或 None。"""
    if not TIME_HINT.search(text):
        return None
    import assistant
    now = _now()
    raw = assistant.chat([{"role": "user", "content": PARSE_PROMPT.format(
        now=f"{now:%Y-%m-%d %H:%M}（週{WEEKDAYS[now.weekday()]}）", text=text)}])
    m = re.search(r"\{.*\}", raw, re.S)
    if not m:
        return None
    try:
        obj = json.loads(m.group(0))
        if not obj.get("is_task") or not obj.get("datetime"):
            return None
        dt = datetime.strptime(obj["datetime"], "%Y-%m-%d %H:%M").replace(tzinfo=TW_TZ)
    except (ValueError, TypeError):
        return None
    if dt <= now:
        return None
    return dt, (obj.get("title") or text).strip()


def add_from_message(text):
    """宗鑫的訊息有時間就建立提醒，回傳要回覆到群組的確認文字；沒有就回 None。"""
    parsed = parse(text)
    if not parsed:
        return None
    event_dt, title = parsed
    remind_dt = max(event_dt - timedelta(minutes=LEAD_MINUTES), _now() + timedelta(minutes=1))
    with _lock:
        data = _load()
        data["alarms"].append({
            "id": uuid.uuid4().hex[:8],
            "event_at": event_dt.isoformat(),
            "remind_at": remind_dt.isoformat(),
            "title": title,
            "created_at": _now().isoformat(timespec="seconds"),
        })
        _save(data)
    return f"⏰ 已設定提醒：{_fmt(remind_dt)} 私訊提醒你\n📌 {_fmt(event_dt)} {title}"


def list_text():
    items = _pending(_load())
    if not items:
        return "目前沒有待提醒的事項"
    lines = [f"{i}. {_fmt(datetime.fromisoformat(a['event_at']))} {a['title']}" for i, a in enumerate(items, 1)]
    return "⏰ 待提醒事項：\n" + "\n".join(lines) + "\n\n取消請輸入：小幫手 取消提醒 編號"


def cancel(index):
    with _lock:
        data = _load()
        items = _pending(data)
        if not 1 <= index <= len(items):
            return f"找不到第 {index} 項，輸入「小幫手 提醒」查看清單"
        target = items[index - 1]
        target["cancelled"] = True
        _save(data)
    return f"🗑️ 已取消：{_fmt(datetime.fromisoformat(target['event_at']))} {target['title']}"


def send_due(push):
    """push(text) -> bool。送出所有到期提醒，回傳送出筆數。"""
    now = _now()
    with _lock:
        data = _load()
        due = [a for a in _pending(data) if datetime.fromisoformat(a["remind_at"]) <= now]
        sent = 0
        for a in due:
            event_dt = datetime.fromisoformat(a["event_at"])
            if push(f"⏰ 提醒：{event_dt:%H:%M} {a['title']}\n（還有 {max(0, int((event_dt - now).total_seconds() // 60))} 分鐘）"):
                a["sent"] = True
                sent += 1
        # 已完成/取消超過 7 天的清掉
        cutoff = now - timedelta(days=7)
        data["alarms"] = [a for a in data["alarms"]
                          if not (a.get("sent") or a.get("cancelled"))
                          or datetime.fromisoformat(a["event_at"]) > cutoff]
        if due:
            _save(data)
    return sent
