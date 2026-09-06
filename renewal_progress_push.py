"""
續保進度推播
------------------------------------------------
流程：
  1. Claude 排程任務（每週一、四 08:30）讀 歸仁續保進度表.xlsx，分析完之後
     在 Drive 的日報表資料夾寫入一個純文字檔 `續保進度_LINE推播_YYYY-MM-DD.txt`，
     第一行固定是 `DATE=YYYY-MM-DD`。
  2. cron-job.org 在 08:35 打 /push-renewal-progress?secret=...
  3. 本模組把那個檔案讀出來，確認是今天的、今天還沒推過，然後推到 LINE 群組。

為什麼要繞這一圈：Claude 排程跑在雲端沙箱，對外連線白名單擋掉了 onrender.com
和 api.line.me，沒辦法直接呼叫 Bot 或 LINE API。但兩邊都連得到 Google Drive，
所以用 Drive 當中繼。

沿用專案既有慣例：
  - 重的 import（googleapiclient 等）全部放在函式內，避免 Render 512MB 記憶體上限
  - 時間一律用明確的 UTC+8
  - 防重複推播的狀態存在 Drive（Render 休眠會清掉本機檔案）
"""

import os
import json
from datetime import datetime, timezone, timedelta

TW_TZ = timezone(timedelta(hours=8))

# 與 drive_reader.DAILY_REPORT_FOLDER_ID 同一個資料夾：
# 服務帳戶對這個資料夾已經有讀寫權限（dashboard_parser 就是往這裡寫 json 的）
REPORT_FOLDER_ID = os.environ.get(
    "RENEWAL_PROGRESS_FOLDER_ID", "1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf"
)
# Claude 排程寫出來的檔名前綴（實際檔名會帶日期，例如 續保進度_LINE推播_2026-09-07.txt）
REPORT_NAME_PREFIX = os.environ.get("RENEWAL_PROGRESS_PREFIX", "續保進度_LINE推播")
STATE_FILENAME = "renewal_progress_pushed.json"

TARGET_GROUP_ID = os.environ.get("REMINDER_GROUP_ID")
LINE_TOKEN = os.environ.get("LINE_TOKEN")

LINE_TEXT_LIMIT = 4900  # LINE 單則上限 5000，留一點餘裕


def _today_str():
    return datetime.now(TW_TZ).strftime("%Y-%m-%d")


def _fetch_latest_report():
    """
    從 Drive 撈最新一份續保進度推播檔。
    回傳 (file_name, date_str, body)；找不到檔案回傳 (None, None, None)。
    date_str 為 None 代表檔案缺少 DATE= 標頭。
    """
    from drive_reader import get_drive_service, download_file

    service = get_drive_service()
    query = (
        f"'{REPORT_FOLDER_ID}' in parents and trashed = false "
        f"and name contains '{REPORT_NAME_PREFIX}'"
    )
    resp = service.files().list(
        q=query,
        orderBy="createdTime desc",
        pageSize=5,
        fields="files(id, name, createdTime)",
    ).execute()
    files = resp.get("files", [])
    if not files:
        return None, None, None

    latest = files[0]
    raw = download_file(latest["id"]).decode("utf-8", errors="replace").strip()
    lines = raw.splitlines()

    if lines and lines[0].startswith("DATE="):
        date_str = lines[0][len("DATE="):].strip()
        body = "\n".join(lines[1:]).strip()
        return latest["name"], date_str, body

    return latest["name"], None, raw


def _push_to_line(text):
    """自己送 LINE，不從 app.py import，避免循環匯入。"""
    import requests

    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json",
    }
    body = {
        "to": TARGET_GROUP_ID,
        "messages": [{"type": "text", "text": text[:LINE_TEXT_LIMIT]}],
    }
    try:
        return requests.post(url, headers=headers, json=body, timeout=10)
    except Exception as e:
        print(f"[ERROR] 續保進度 LINE 推播失敗: {e}")
        return None


def run_progress_push(force=False):
    """
    回傳 (ok: bool, message: str)。
    force=True 會跳過「日期必須是今天」和「今天沒推過」兩道檢查，用來手動測試。
    """
    if not LINE_TOKEN:
        return False, "LINE_TOKEN 未設定"
    if not TARGET_GROUP_ID:
        return False, "REMINDER_GROUP_ID 未設定"

    file_name, date_str, body = _fetch_latest_report()
    if body is None:
        return False, f"Drive 資料夾 {REPORT_FOLDER_ID} 裡找不到 {REPORT_NAME_PREFIX}*"
    if not body.strip():
        return False, f"{file_name} 內容是空的，不推播"

    today = _today_str()

    if not force:
        if date_str is None:
            return False, f"{file_name} 缺少 DATE= 標頭，無法確認是不是今天的，略過"
        if date_str != today:
            return False, f"{file_name} 是 {date_str} 的報告，今天是 {today}，略過"

        from drive_json_store import load_json_from_drive

        state = load_json_from_drive(REPORT_FOLDER_ID, STATE_FILENAME) or {}
        if state.get("last_pushed_date") == today:
            return False, f"今天（{today}）已經推播過了，略過"

    resp = _push_to_line(body)
    if resp is None:
        return False, "LINE 推播沒有回應（連線失敗或逾時）"
    if resp.status_code != 200:
        return False, f"LINE 推播失敗 HTTP {resp.status_code}: {resp.text[:200]}"

    if not force:
        from drive_json_store import save_json_to_drive

        save_json_to_drive(
            REPORT_FOLDER_ID,
            STATE_FILENAME,
            {
                "last_pushed_date": today,
                "pushed_at": datetime.now(TW_TZ).isoformat(timespec="seconds"),
                "source_file": file_name,
            },
        )

    return True, f"已推播 {file_name}（{len(body)} 字）到群組"
