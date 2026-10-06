"""
觸發任務「建道學長升主任倒數」— LINE 即時回報
------------------------------------------------
宗鑫在 LINE 傳：
  建道+1 / 建道+2 → 記錄成交（立即反映在儀表板）
  建道-1          → 更正
  建道進度        → 查目前進度
日報表進來後，以日報表為準：只有「晚於最新日報表建立時間」的回報才會被加上去，
所以白天回報、晚上日報表確認後自動作廢，不會重複計算。
回報紀錄存在 Drive：promo_jiandao_manual.json（日報表資料夾）
"""
import re
import threading
from datetime import datetime, timezone, timedelta

TPE = timezone(timedelta(hours=8))
MANUAL_FILENAME = "promo_jiandao_manual.json"
_PAT = re.compile(r"^建道\s*(?:([+＋\-－])\s*(\d{0,2})\s*台?|進度)$")
_cache = {"entries": None}
_lock = threading.Lock()


def _now():
    return datetime.now(TPE).replace(tzinfo=None).isoformat(timespec="seconds")


def load(force=False):
    if _cache["entries"] is None or force:
        from drive_json_store import load_json_from_drive
        from drive_reader import DAILY_REPORT_FOLDER_ID
        data = load_json_from_drive(DAILY_REPORT_FOLDER_ID, MANUAL_FILENAME) or {}
        _cache["entries"] = list(data.get("entries", []))
    return list(_cache["entries"])


def _add(delta):
    from drive_json_store import save_json_to_drive
    from drive_reader import DAILY_REPORT_FOLDER_ID
    with _lock:
        entries = load()
        entries.append({"ts": _now(), "delta": delta})
        save_json_to_drive(DAILY_REPORT_FOLDER_ID, MANUAL_FILENAME, {"entries": entries})
        _cache["entries"] = entries


def match(text):
    return bool(_PAT.match(text.replace(" ", "")))


def _report_status():
    """依最新日報表算出的台數 + 晚於日報表的 LINE 回報。"""
    import os, json
    from dashboard_parser import DATA_FILE
    gr = None
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, encoding="utf-8") as f:
            gr = json.load(f)
    else:
        from drive_json_store import restore_cache
        gr = restore_cache(DATA_FILE)
    p = (gr or {}).get("promo") or {}
    if not p.get("start"):
        return None
    pts = {d: v for d, v in (p.get("points") or {}).items() if p["start"] <= d <= p["end"]}
    last_d = max(pts) if pts else p["start"]
    adj = sum(a[1] for a in p.get("adjust", []) if a[0] <= last_d)
    report_val = p["base"] + adj + pts.get(last_d, 0)
    rc = p.get("report_created") or (last_d + "T17:00:00")
    manual = sum(e["delta"] for e in load() if e["ts"] > rc)
    return {"report": report_val, "manual": manual, "cur": report_val + manual,
            "target": p["target"], "last_d": last_d}


def handle(text):
    m = _PAT.match(text.replace(" ", ""))
    sign, num = m.group(1), m.group(2)
    msg = []
    if sign:
        n = int(num) if num else 1
        if n == 0 or n > 5:
            return "數字怪怪的，一次請回報 1～5 台，例如：建道+1"
        delta = n if sign in "+＋" else -n
        try:
            _add(delta)
        except Exception as e:
            return f"記錄失敗，請稍後再試（{str(e)[:60]}）"
        msg.append(f"✅ 已記錄 建道 {'+' if delta > 0 else ''}{delta}（{_now()[5:16].replace('-', '/').replace('T', ' ')}）")
    try:
        s = _report_status()
    except Exception as e:
        s = None
        print(f"[promo] status error: {e}")
    if not s:
        msg.append("目前還讀不到儀表板資料，回報已存好，晚點會反映。")
        return "\n".join(msg)
    left = max(0, s["target"] - s["cur"])
    md = f"{int(s['last_d'][5:7])}/{int(s['last_d'][8:10])}"
    detail = f"（日報表 {md}：{s['report']} 台" + (f"，LINE 回報 {s['manual']:+d}" if s["manual"] else "") + "）"
    msg.append(f"🎯 建道學長目前 {s['cur']} 台{detail}")
    msg.append("🎉 已達標！恭喜升主任！" if left == 0 else f"距升主任還差 {left} 台 🔥")
    return "\n".join(msg)
