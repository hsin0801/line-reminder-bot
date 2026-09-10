"""
LINE 月額度守門員
------------------------------------------------
為什麼需要這個模組：

LINE 的訊息則數是「以發送對象人數」計算的——推播到一個 9 人群組算 9 則，
不是 1 則。免費（輕用量）方案每月 200 則，換算下來一個月只夠推 22 次群組訊息。

2026-09-10 就是因為額度用完，續保進度推播整天卡在
`HTTP 429 You have reached your monthly limit.`，而且是靜默失敗——
不看 Render log 根本不知道。

這個模組做三件事：
  1. 把「這個月用掉幾則」記在 Drive（Render 免費方案本機硬碟休眠就清空，
     存本機等於沒存，這是 renewal_reminder.py 原本踩過的坑）
  2. 送出前先問「還夠不夠」，不夠就不送，並回報剩餘額度
  3. 分優先權：低優先的每小時提醒不能把高優先的週一/四續保進度推播吃掉

群組人數用 LINE API 實際查，每月快取一次；查不到就用 DEFAULT_GROUP_SIZE 保守估。

沿用專案既有慣例：重的 import 放函式內（Render 512MB 上限）、時間一律 UTC+8。
"""

import os
import json
from datetime import datetime, timezone, timedelta

TW_TZ = timezone(timedelta(hours=8))

# 與 renewal_progress_push.REPORT_FOLDER_ID 同一個資料夾
FOLDER_ID = os.environ.get(
    "RENEWAL_PROGRESS_FOLDER_ID", "1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf"
)
QUOTA_FILENAME = "line_quota_state.json"

# 每月額度。輕用量(免費)=200、中用量=3000、高用量=6000。
# 升方案之後記得把 Render 的 LINE_MONTHLY_LIMIT 環境變數一起改。
MONTHLY_LIMIT = int(os.environ.get("LINE_MONTHLY_LIMIT", "200"))

# 保留給高優先推播（週一/四續保進度）的額度。
# 低優先（每小時 @提醒、每日行程提醒）只能用到 MONTHLY_LIMIT - RESERVE。
RESERVE_FOR_HIGH = int(os.environ.get("LINE_QUOTA_RESERVE", "100"))

# 查不到群組人數時的保守預設值（寧可高估，高估只會提早擋下來，不會超發）
DEFAULT_GROUP_SIZE = int(os.environ.get("LINE_GROUP_SIZE", "10"))

LINE_TOKEN = os.environ.get("LINE_TOKEN")


def _now():
    return datetime.now(TW_TZ)


def _month_key(dt=None):
    return (dt or _now()).strftime("%Y-%m")


def _load():
    from drive_json_store import load_json_from_drive

    try:
        return load_json_from_drive(FOLDER_ID, QUOTA_FILENAME) or {}
    except Exception as e:
        print(f"[WARN] 讀取 {QUOTA_FILENAME} 失敗，本次不擋: {e}")
        return {}


def _save(state):
    """寫入失敗不能讓呼叫端爆掉——訊息可能已經送出去了。"""
    from drive_json_store import save_json_to_drive

    try:
        save_json_to_drive(FOLDER_ID, QUOTA_FILENAME, state)
        return True
    except Exception as e:
        print(f"[ERROR] 寫入 {QUOTA_FILENAME} 失敗，額度計數會不準: {e}")
        return False


def fetch_group_size(group_id):
    """
    問 LINE 這個群組有幾個人。查不到回 None。
    注意：只算「有加 Bot 好友」的成員，實際計費人數以 LINE 後台為準，
    所以這個數字是下限，我們另外用 DEFAULT_GROUP_SIZE 兜底。
    """
    if not LINE_TOKEN or not group_id:
        return None
    import requests

    url = f"https://api.line.me/v2/bot/group/{group_id}/members/count"
    try:
        resp = requests.get(
            url, headers={"Authorization": f"Bearer {LINE_TOKEN}"}, timeout=10
        )
        if resp.status_code == 200:
            return int(resp.json().get("count", 0)) or None
        print(f"[WARN] 查群組人數失敗 HTTP {resp.status_code}: {resp.text[:150]}")
    except Exception as e:
        print(f"[WARN] 查群組人數例外: {e}")
    return None


def _ensure_month(state, group_id):
    """跨月自動歸零，順便重新查一次群組人數。"""
    mk = _month_key()
    if state.get("month") != mk:
        size = fetch_group_size(group_id) or DEFAULT_GROUP_SIZE
        state = {
            "month": mk,
            "used": 0,
            "group_size": size,
            "limit": MONTHLY_LIMIT,
            "history": [],
        }
        print(f"[QUOTA] 進入新月份 {mk}，額度歸零，群組人數 {size}")
    elif not state.get("group_size"):
        state["group_size"] = fetch_group_size(group_id) or DEFAULT_GROUP_SIZE
    return state


def check(group_id, pushes=1, priority="normal"):
    """
    問「現在送 pushes 次群組推播，額度夠不夠」。
    回傳 (ok: bool, info: dict)。不實際扣款，扣款請呼叫 commit()。

    priority="high"   → 可以用到整條 MONTHLY_LIMIT
    priority="normal" → 只能用到 MONTHLY_LIMIT - RESERVE_FOR_HIGH
    """
    state = _ensure_month(_load(), group_id)
    size = state.get("group_size") or DEFAULT_GROUP_SIZE
    used = int(state.get("used", 0))
    cost = pushes * size

    ceiling = MONTHLY_LIMIT if priority == "high" else max(0, MONTHLY_LIMIT - RESERVE_FOR_HIGH)
    ok = (used + cost) <= ceiling

    info = {
        "state": state,
        "group_size": size,
        "used": used,
        "cost": cost,
        "ceiling": ceiling,
        "limit": MONTHLY_LIMIT,
        "remaining": max(0, ceiling - used),
        "remaining_pushes": max(0, (ceiling - used) // size) if size else 0,
    }
    if not ok:
        print(
            f"[QUOTA] 擋下：本月已用 {used}/{MONTHLY_LIMIT} 則，"
            f"本次需 {cost} 則（{pushes}次 × {size}人），"
            f"{priority} 上限 {ceiling} 則"
        )
    return ok, info


def commit(info, pushes_sent, label=""):
    """訊息確實送出去之後才呼叫，把用量寫回 Drive。"""
    state = info["state"]
    size = info["group_size"]
    spent = pushes_sent * size
    state["used"] = int(state.get("used", 0)) + spent
    state["limit"] = MONTHLY_LIMIT
    state.setdefault("history", []).append(
        {
            "at": _now().isoformat(timespec="seconds"),
            "label": label,
            "pushes": pushes_sent,
            "cost": spent,
            "used_after": state["used"],
        }
    )
    # history 只留最近 60 筆，免得檔案無限長大
    state["history"] = state["history"][-60:]
    _save(state)
    print(
        f"[QUOTA] {label} 送出 {pushes_sent} 次 × {size} 人 = {spent} 則，"
        f"本月累計 {state['used']}/{MONTHLY_LIMIT}"
    )
    return state["used"]


def status(group_id):
    """給人看的目前用量，可以掛一條路由出來查。"""
    state = _ensure_month(_load(), group_id)
    size = state.get("group_size") or DEFAULT_GROUP_SIZE
    used = int(state.get("used", 0))
    normal_ceiling = max(0, MONTHLY_LIMIT - RESERVE_FOR_HIGH)
    return {
        "month": state.get("month"),
        "group_size": size,
        "used": used,
        "limit": MONTHLY_LIMIT,
        "remaining": max(0, MONTHLY_LIMIT - used),
        "remaining_pushes_high": max(0, (MONTHLY_LIMIT - used) // size) if size else 0,
        "remaining_pushes_normal": max(0, (normal_ceiling - used) // size) if size else 0,
        "reserve_for_high": RESERVE_FOR_HIGH,
    }
