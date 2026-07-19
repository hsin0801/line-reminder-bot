import os
import json
import requests
import random
import time
import importlib  # ─── 【新增】用於強迫刷新模組的工具 ───
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
except ImportError as e:
    import traceback
    print(f"[WARNING] dashboard_routes 載入失敗: {e}")
    print(traceback.format_exc())

LINE_TOKEN = os.environ.get("LINE_TOKEN")
BASE_URL = "https://line-reminder-bot-gj9p.onrender.com/img"
HSIN_USER_ID = "U272a3c6b1f3d10a3677769cb4f73fe1d"

# 記憶體計數器（處理群組內部的次數計數）
stock_count = {}
po_count = {}


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
                "originalContentUrl": f"{BASE_URL}/july.jpg{cache_buster}",
                "previewImageUrl": f"{BASE_URL}/july.jpg{cache_buster}"
            }])

        elif text == "SP":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📄 SP 活動資料：\nhttps://drive.google.com/file/d/1MDr9BQAZURlVlUYMbSjARvWRmn7exIp3/view?usp=sharing"
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


# ── 5. 業績速報推播路由（中午12點保底 + 儲存至檔案防 Render 失憶 + 模組強迫重整機制） ────
@app.route("/push-speed-report", methods=["GET"])
def push_speed_report():
    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401

    # 取得台灣當前時間
    tw_now = datetime.now(timezone(timedelta(hours=8)))
    today_str = tw_now.strftime("%Y-%m-%d")
    current_hour = tw_now.hour

    # 從檔案撈取上一次的紀錄（徹底搞定重啟失憶）
    last_pushed_report_date, has_pushed_fallback_today = get_report_status_from_file()

    try:
        # ──── 【核心優化】強迫 Python 重新讀取 drive_reader，擊碎模組清單快取 ────
        import drive_reader
        importlib.reload(drive_reader)
        
        report = drive_reader.get_speed_report()
        report_date = report.get("date", "")  # 例如 "20260719"

        if not report_date:
            print("[WARN] 無化取得速報日期")
            return "Report date missing", 200

        # 核心比對：最新的速報日期跟檔案裡的不一樣，代表 Drive 裡長出全新檔名的資料了！
        if report_date != last_pushed_report_date:
            message = drive_reader.format_speed_report_message(report)
            full_message = f"🔥 【最新業績速報更新！】\n\n{message}"
            
            resp = push_message(HSIN_USER_ID, [{"type": "text", "text": full_message}])
            
            if resp and resp.status_code == 200:
                save_report_status_to_file(report_date=report_date)
                return f"New report pushed: {report_date}", 200
            else:
                return "Push failed", 500

        # 【中午 12 點保底機制】
        # 如果抓到的資料沒更新，且時間已經到了中午 12 點（或之後）
        if current_hour >= 12:
            if has_pushed_fallback_today != today_str:
                fallback_msg = f"🤖 報告主管：\n目前已過中午 {current_hour}:00，但後台的業績速報今天尚未更新新資料唷！\n\n（目前最新仍為 {report_date} 的數據）"
                resp = push_message(HSIN_USER_ID, [{"type": "text", "text": fallback_msg}])
                
                if resp and resp.status_code == 200:
                    save_report_status_to_file(fallback_date=today_str)
                    return "Fallback message pushed", 200

        print(f"[WAIT] 速報日期 {report_date} 已推播過，且尚未到中午，或今日已發過保底。")
        return "No new update template", 200

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


# ── 7. 固定提醒與測試路由 ──────────────────────────────────────
@app.route("/remind/<key>", methods=["GET"])
def remind(key):
    with open("reminders.json", "r", encoding="utf-8") as f:
        config = json.load(f)
    groups = config["groups"]
    reminders = config["reminders"]
    if key not in reminders:
        return "Not found", 404
    message = reminders[key]["message"]
    for group_key in groups:
        push_url = "https://api.line.me/v2/bot/message/push"
        headers = {
            "Authorization": f"Bearer {LINE_TOKEN}",
            "Content-Type": "application/json"
        }
        body = {"to": groups[group_key], "messages": [{"type": "text", "text": message}]}
        requests.post(push_url, headers=headers, json=body)
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
