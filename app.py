import os
import json
import requests
import random
import time
import importlib
from datetime import datetime, timezone, timedelta
from flask import Flask, request, send_from_directory
from renewal_reminder import run_reminder, mark_replied

app = Flask(__name__)

# ──── 1. 先註冊所有 Blueprint，確保路由在啟動前完全載入 ────
try:
    from dashboard_routes import dashboard_bp, yongkang_bp, faren_bp, combined_bp
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(yongkang_bp)
    app.register_blueprint(faren_bp)
    app.register_blueprint(combined_bp)
except ImportError:
    print("[WARNING] dashboard_routes 載入失敗，請確認檔案是否存在")

LINE_TOKEN = os.environ.get("LINE_TOKEN")
BASE_URL = "https://line-reminder-bot-gj9p.onrender.com/img"
HSIN_USER_ID = "U272a3c6b1f3d10a3677769cb4f73fe1d"

# 記憶體計數器（處理群組內部的次數計數）
stock_count = {}
po_count = {}
_speed_report_pushed_date = None  # 記憶體防重複推播


# ──── 2. 讀寫 reminders.json 的工具函式（防 Render 休眠失憶） ────
def get_report_status_from_file():
    """從 json 檔案讀取上一次推播的狀態"""
    filename = "reminders.json"
    if not os.path.exists(filename):
        return None, None
    try:
        with open(filename, "r", encoding="utf-8") as f:
            config = json.load(f)
        return config.get("last_speed_report_date"), config.get("last_fallback_date")
    except Exception as e:
        print(f"[ERROR] 讀取狀態檔案失敗: {e}")
        return None, None

def save_report_status_to_file(report_date=None, fallback_date=None):
    """將推播狀態寫入 json 檔案，防止重啟遺失"""
    filename = "reminders.json"
    try:
        config = {}
        if os.path.exists(filename):
            with open(filename, "r", encoding="utf-8") as f:
                config = json.load(f)

        if report_date is not None:
            config["last_speed_report_date"] = report_date
        if fallback_date is not None:
            config["last_fallback_date"] = fallback_date

        with open(filename, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"[DISK] 成功將狀態寫入檔案。Report: {report_date}, Fallback: {fallback_date}")
    except Exception as e:
        print(f"[ERROR] 寫入狀態檔案失敗: {e}")


# ──── 3. LINE 基本傳送函式（加入逾時機制） ────
def reply_message(reply_token, messages):
    url = "https://api.line.me/v2/bot/message/reply"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    body = {"replyToken": reply_token, "messages": messages}
    try:
        requests.post(url, headers=headers, json=body, timeout=5)
    except requests.exceptions.Timeout:
        print("[TIMEOUT] LINE reply 逾時")

def push_message(to, messages):
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    body = {"to": to, "messages": messages}
    try:
        return requests.post(url, headers=headers, json=body, timeout=5)
    except Exception as e:
        print(f"[ERROR] Push message failed: {e}")
        return None


# ──── 4. LINE Webhook 訊息主路由 ────
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()
    events = data.get("events", [])

    for event in events:
        if event.get("type") != "message" or event.get("message", {}).get("type") != "text":
            continue

        text = event["message"]["text"].strip()
        reply_token = event["replyToken"]
        source = event.get("source", {})
        user_id = source.get("userId", "unknown")
        group_id = source.get("groupId", "")

        # 回覆偵測（加入 Timeout，防止抓 Profile 卡死 Webhook）
        if group_id == os.environ.get("REMINDER_GROUP_ID"):
            display_name = ""
            if user_id != "unknown":
                try:
                    profile_url = f"https://api.line.me/v2/bot/group/{group_id}/member/{user_id}"
                    headers_profile = {"Authorization": f"Bearer {LINE_TOKEN}"}
                    resp = requests.get(profile_url, headers=headers_profile, timeout=3)
                    if resp.status_code == 200:
                        display_name = resp.json().get("displayName", "")
                except Exception as e:
                    print(f"[ERROR] Get Profile Failed: {e}")
            if display_name:
                mark_replied(display_name)

        # 關鍵字指令區（圖片網址後加上 ?t=時間戳記，防快取舊圖）
        cache_buster = f"?t={int(time.time())}"

        if text == "內促":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/september.png{cache_buster}",
                "previewImageUrl": f"{BASE_URL}/september.png{cache_buster}"
            }])

        elif text == "SP":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📄 SP 活動資料：\nhttps://drive.google.com/file/d/12NhD5qABCAccfPJZnKn-KzugfJfcgpkz/view?usp=sharing"
            }])

        elif text == "配件":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📊 配件資料：\nhttps://docs.google.com/spreadsheets/d/1ck3utTd1TIAY2ZtiyCrKk5mjitQBYehZ/edit?usp=sharing"
            }])

        elif text == "組合價":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/combination.png{cache_buster}",
                "previewImageUrl": f"{BASE_URL}/combination.png{cache_buster}"
            }])

        elif text == "業績儀表板":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📊 業績儀表板：\nhttps://line-reminder-bot-gj9p.onrender.com/warroom/"
            }])

        elif text == "紀柏州":
            quote_token = event.get("message", {}).get("quoteToken", "")
            po_count[user_id] = po_count.get(user_id, 0) + 1

            if po_count[user_id] >= 3:
                po_count[user_id] = 0
                display_name = "你"
                if group_id and user_id != "unknown":
                    try:
                        profile_url = f"https://api.line.me/v2/bot/group/{group_id}/member/{user_id}"
                        headers_p = {"Authorization": f"Bearer {LINE_TOKEN}"}
                        resp = requests.get(profile_url, headers=headers_p, timeout=2)
                        if resp.status_code == 200:
                            display_name = resp.json().get("displayName", "你")
                    except:
                        pass

                mention_text = f"@{display_name}"
                full_text = f"{mention_text} 你不要那麼愛我 明天14:00來永康找我開會 ❤️"
                reply_msg = {
                    "type": "text",
                    "text": full_text,
                    "mention": {
                        "mentionees": [{
                            "index": 0,
                            "length": len(mention_text),
                            "type": "user",
                            "userId": user_id
                        }]
                    }
                }
                if quote_token:
                    reply_msg["quoteToken"] = quote_token
                reply_message(reply_token, [reply_msg])
            else:
                po_images = ["po.png", "po2.png", "po3.png", "po4.jpg"]
                img = random.choice(po_images)
                reply_message(reply_token, [{
                    "type": "image",
                    "originalContentUrl": f"{BASE_URL}/{img}{cache_buster}",
                    "previewImageUrl": f"{BASE_URL}/{img}{cache_buster}"
                }])

        elif text in ["陳建道", "陳星佑", "歐陽", "午安猴", "張姉瑀", "林定緯"]:
            mapping = {
                "陳建道": ["dao.jpg", "dao2.jpg", "dao3.jpg", "dao4.jpg", "dao5.jpg", "dao6.jpg", "dao7.jpg"],
                "陳星佑": ["chen.jpg", "chen2.jpg", "chen3.jpg", "chen4.jpg"],
                "歐陽": ["OY.jpg", "OY2.jpg"],
                "午安猴": ["hao.jpg", "hao2.jpg", "hao3.jpg"],
                "張姉瑀": ["fish.jpg"],
                "林定緯": ["ding.jpg"]
            }
            img = random.choice(mapping[text])
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/{img}{cache_buster}",
                "previewImageUrl": f"{BASE_URL}/{img}{cache_buster}"
            }])

        elif text == "劉宗鑫":
            quote_token = event.get("message", {}).get("quoteToken", "")
            reply_msg = {"type": "text", "text": "賴翔德～我是不會屈服的"}
            if quote_token:
                reply_msg["quoteToken"] = quote_token
            reply_message(reply_token, [reply_msg])

        elif text == "條件":
            reply_message(reply_token, [{
                "type": "text",
                "text": "https://honda-bonus-calculator-kmqxbzbyowyzwbelmz82lb.streamlit.app/#2026-5-honda"
            }])

        elif text == "接龍":
            reply_message(reply_token, [{
                "type": "text",
                "text": "宗鑫 \n定緯 \n適緯 \n建道 \n星佑 \n姉瑀 \n文智 \n明憬 "
            }])

        elif text.startswith("小幫手"):
            question = text[3:].strip()
            if not question:
                reply_message(reply_token, [{"type": "text", "text": "請在「小幫手」後面輸入你的問題！"}])
            else:
                try:
                    groq_url = "https://api.groq.com/openai/v1/chat/completions"
                    groq_headers = {
                        "Authorization": f"Bearer {os.environ.get('GROQ_API_KEY')}",
                        "Content-Type": "application/json"
                    }
                    groq_body = {
                        "model": "llama-3.1-8b-instant",
                        "messages": [
                            {"role": "system", "content": "你是一個汽車業務團隊的AI小幫手，專門協助回答業務銷售相關問題。請用繁體中文簡潔搞笑地回答。"},
                            {"role": "user", "content": question}
                        ]
                    }
                    resp = requests.post(groq_url, headers=groq_headers, json=groq_body, timeout=6)
                    answer = resp.json()["choices"][0]["message"]["content"]
                    reply_message(reply_token, [{"type": "text", "text": f"🤖 {answer}"}])
                except Exception as e:
                    reply_message(reply_token, [{"type": "text", "text": f"小幫手開小差了，等等再試！"}])

        elif text == "推薦股票":
            stock_count[user_id] = stock_count.get(user_id, 0) + 1
            if stock_count[user_id] >= 3:
                stock_count[user_id] = 0
                reply_message(reply_token, [{"type": "text", "text": "戒斷當沖 是你唯一選擇"}])
            else:
                stocks = [("2330", "台積電"), ("2317", "鴻海"), ("2454", "聯發科"), ("2382", "廣達")]
                code, name = random.choice(stocks)
                reply_message(reply_token, [{"type": "text", "text": f"📈 今日推薦股票\n\n【{code} {name}】\n\n⚠️ 僅供娛樂，不構成投資建議！"}])

    return "OK", 200


# ── 5. 業績速報推播路由 ────
@app.route("/push-speed-report", methods=["GET"])
def push_speed_report():
    global _speed_report_pushed_date

    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401

    # 取得台灣當前時間
    tw_now = datetime.now(timezone(timedelta(hours=8)))
    today_str = tw_now.strftime("%Y-%m-%d")
    current_hour = tw_now.hour

    # ── 記憶體層防重複：同一個 Render session 已發過今天，直接跳過 ──
    if _speed_report_pushed_date == today_str:
        print(f"[SKIP] 今天 {today_str} 業績速報已推播（記憶體）")
        return "OK", 200

    # 從檔案撈取上一次的紀錄
    last_pushed_report_date, has_pushed_fallback_today = get_report_status_from_file()

    # ── 檔案層防重複：冷啟動後也能跳過 ──
    if last_pushed_report_date == today_str:
        _speed_report_pushed_date = today_str
        print(f"[SKIP] 今天 {today_str} 業績速報已推播（檔案）")
        return "OK", 200

    try:
        # 強迫 Python 重新讀取 drive_reader，擊碎模組清單快取
        import drive_reader
        importlib.reload(drive_reader)

        report = drive_reader.get_speed_report()
        report_date = report.get("date", "")  # 例如 "20260716"

        if not report_date:
            print("[WARN] 無法取得速報日期")
            return "Report date missing", 200

        # 檢查這份速報日期是否太舊（超過 4 天以上）
        try:
            parsed_report_date = datetime.strptime(report_date, "%Y%m%d").date()
            tw_today = tw_now.date()

            if (tw_today - parsed_report_date).days > 4:
                print(f"[SKIP] 速報日期 {report_date} 為舊資訊，不執行日常推播。")

                if current_hour >= 12 and has_pushed_fallback_today != today_str:
                    fallback_msg = f"🤖 報告主管：\n目前已過中午 {current_hour}:00，但後台的業績速報今天尚未更新新資料唷！\n\n（目前最新仍為 {report_date} 的數據）"
                    resp = push_message(HSIN_USER_ID, [{"type": "text", "text": fallback_msg}])
                    if resp and resp.status_code == 200:
                        save_report_status_to_file(fallback_date=today_str)
                        return "Fallback message pushed due to old data date", 200

                return f"Skipped old report: {report_date}", 200

        except Exception as e:
            print(f"[WARN] 解析速報日期與今日比對失敗: {e}")

        # 核心比對：速報有新資料就推播
        if report_date != last_pushed_report_date:
            message = drive_reader.format_speed_report_message(report)
            full_message = f"🔥 【最新業績速報更新！】\n\n{message}"

            resp = push_message(HSIN_USER_ID, [{"type": "text", "text": full_message}])

            if resp and resp.status_code == 200:
                _speed_report_pushed_date = today_str          # 標記記憶體
                save_report_status_to_file(report_date=today_str)  # 存今天日期防重複
                return f"New report pushed: {report_date}", 200
            else:
                return "Push failed", 500

        # 保底機制：資料沒更新但時間已到中午 12 點
        if current_hour >= 12:
            if has_pushed_fallback_today != today_str:
                fallback_msg = f"🤖 報告主管：\n目前已過中午 {current_hour}:00，但後台的業績速報今天尚未更新新資料唷！\n\n（目前最新仍為 {report_date} 的數據）"
                resp = push_message(HSIN_USER_ID, [{"type": "text", "text": fallback_msg}])

                if resp and resp.status_code == 200:
                    save_report_status_to_file(fallback_date=today_str)
                    return "Fallback message pushed", 200

        print(f"[WAIT] 速報日期 {report_date} 已推播過，或今日已發過保底。")
        return "No new update", 200

    except Exception as e:
        import traceback
        print(f"[ERROR] push_speed_report:\n{traceback.format_exc()}")
        return f"Error: {e}", 500


# ── 6. 續保提醒觸發路由 ──────────────────────────────────
@app.route("/run-renewal-reminder", methods=["GET"])
def run_renewal_reminder():
    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401
    try:
        run_reminder()
        return "OK", 200
    except Exception as e:
        print(f"[ERROR] run_reminder: {e}")
        return f"Error: {e}", 500


# ── 續保進度推播（Claude 排程寫 Drive → 這裡讀出來推 LINE）──
# Claude 排程每週一、四 08:30 分析續保進度表，把 LINE 版報告寫到 Drive；
# cron-job.org 每週一、四 08:35 打這條路由，把它推到群組。
# 加 &force=1 可跳過「必須是今天」與「今天沒推過」的檢查，用來手動測試。
@app.route("/push-renewal-progress", methods=["GET"])
def push_renewal_progress():
    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401
    force = request.args.get("force", "") == "1"
    try:
        from renewal_progress_push import run_progress_push
        ok, msg = run_progress_push(force=force)
        print(f"[RENEWAL-PROGRESS] {'OK' if ok else 'SKIP'} - {msg}")
        # 略過（報告還沒寫、不是今天的、今天推過了）不是錯誤，
        # 回 200 避免 cron-job.org 一直發失敗通知。
        return msg, 200
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[ERROR] push_renewal_progress: {e}")
        return f"Error: {e}", 500


# ── 7. 固定提醒與測試路由 ──────────────────────────────────────

# ── 歸仁日報表 vs 週邊指標 差異比對 ──────────────────────────
@app.route("/run-kpi-check", methods=["GET"])
def run_kpi_check_route():
    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401
    try:
        from guiren_kpi_checker import run_kpi_check
        from drive_reader import get_drive_service
        run_kpi_check(get_drive_service())
        return "OK", 200
    except Exception as e:
        print(f"[ERROR] run_kpi_check: {e}")
        return f"Error: {e}", 500

@app.route("/remind/<key>", methods=["GET"])
def remind(key):
    with open("reminders.json", "r", encoding="utf-8") as f:
        config = json.load(f)
    groups = config["groups"]
    reminders = config["reminders"]
    if key not in reminders:
        return "Not found", 404

    reminder = reminders[key]
    message = reminder["message"]
    mention_names = reminder.get("mentions", [])

    msg = {"type": "text", "text": message}

    if mention_names:
        member_ids = json.loads(os.environ.get("MEMBER_USER_IDS", "{}"))
        mentionees = []
        for name in mention_names:
            at = f"@{name}"
            idx = message.find(at)
            if idx != -1 and name in member_ids:
                mentionees.append({
                    "index": idx,
                    "length": len(at),
                    "type": "user",
                    "userId": member_ids[name]
                })
        if mentionees:
            msg["mention"] = {"mentionees": mentionees}

    target_groups = reminder.get("groups", list(groups.keys()))
    for group_key in target_groups:
        if group_key in groups:
            push_message(groups[group_key], [msg])

    return "OK", 200

@app.route("/test-drive", methods=["GET"])
def test_drive():
    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401
    import traceback
    from drive_reader import get_speed_report, get_daily_report
    result = {}
    try:
        result["speed_report"] = get_speed_report()
    except Exception as e:
        result["speed_error"] = traceback.format_exc()
    try:
        result["daily_report"] = get_daily_report()
    except Exception as e:
        result["daily_error"] = traceback.format_exc()
    return json.dumps(result, ensure_ascii=False, indent=2), 200


# ── 8. 基本路由 ──────────────────────────────────────────
@app.route("/", methods=["GET"])
def index():
    return "LINE Bot is running!", 200

@app.route("/img/<filename>")
def serve_image(filename):
    return send_from_directory(".", filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

from dashboard_routes import dashboard_bp, yongkang_bp, faren_bp, combined_bp, warroom_bp, _guiren_kpi_bp
app.register_blueprint(warroom_bp)
app.register_blueprint(_guiren_kpi_bp)
