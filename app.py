import os
import json
import requests
import random
from datetime import datetime
from flask import Flask, request, send_from_directory
from renewal_reminder import run_reminder, mark_replied

app = Flask(__name__)

LINE_TOKEN = os.environ.get("LINE_TOKEN")
BASE_URL = "https://line-reminder-bot-gj9p.onrender.com/img"
HSIN_USER_ID = "U272a3c6b1f3d10a3677769cb4f73fe1d"
stock_count = {}
po_count = {}

def reply_message(reply_token, messages):
    url = "https://api.line.me/v2/bot/message/reply"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    body = {
        "replyToken": reply_token,
        "messages": messages
    }
    requests.post(url, headers=headers, json=body)

def push_message(to, messages):
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    body = {"to": to, "messages": messages}
    return requests.post(url, headers=headers, json=body)

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()
    events = data.get("events", [])

    for event in events:
        if event.get("type") != "message":
            continue
        if event.get("message", {}).get("type") != "text":
            continue

        text = event["message"]["text"].strip()
        reply_token = event["replyToken"]
        source = event.get("source", {})
        user_id = source.get("userId", "unknown")
        group_id = source.get("groupId", "")

        # 回覆偵測
        if group_id == os.environ.get("REMINDER_GROUP_ID"):
            display_name = ""
            if user_id != "unknown":
                try:
                    profile_url = f"https://api.line.me/v2/bot/group/{group_id}/member/{user_id}"
                    headers_profile = {"Authorization": f"Bearer {LINE_TOKEN}"}
                    resp = requests.get(profile_url, headers=headers_profile)
                    if resp.status_code == 200:
                        display_name = resp.json().get("displayName", "")
                except:
                    pass
            if display_name:
                mark_replied(display_name)

        if text == "內促":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/june.jpg",
                "previewImageUrl": f"{BASE_URL}/june.jpg"
            }])

        elif text == "SP":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📄 SP 活動資料：\nhttps://drive.google.com/file/d/1EZCnZPdT6fh8FwW_VsAbUtIY_aQUuuO5/view?usp=sharing"
            }])

        elif text == "配件":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📊 配件資料：\nhttps://docs.google.com/spreadsheets/d/1ck3utTd1TIAY2ZtiyCrKk5mjitQBYehZ/edit?usp=sharing&ouid=109189035277985438460&rtpof=true&sd=true"
            }])

        elif text == "組合價":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/combination.jpg",
                "previewImageUrl": f"{BASE_URL}/combination.jpg"
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
                        resp = requests.get(profile_url, headers=headers_p)
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
                    "originalContentUrl": f"{BASE_URL}/{img}",
                    "previewImageUrl": f"{BASE_URL}/{img}"
                }])

        elif text == "陳建道":
            dao_images = ["dao.jpg", "dao2.jpg", "dao3.jpg", "dao4.jpg", "dao5.jpg", "dao6.jpg", "dao7.jpg"]
            img = random.choice(dao_images)
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/{img}",
                "previewImageUrl": f"{BASE_URL}/{img}"
            }])

        elif text == "陳星佑":
            chen_images = ["chen.jpg", "chen2.jpg", "chen3.jpg", "chen4.jpg"]
            img = random.choice(chen_images)
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/{img}",
                "previewImageUrl": f"{BASE_URL}/{img}"
            }])

        elif text == "歐陽":
            oy_images = ["OY.jpg", "OY2.jpg"]
            img = random.choice(oy_images)
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/{img}",
                "previewImageUrl": f"{BASE_URL}/{img}"
            }])

        elif text == "劉宗鑫":
            quote_token = event.get("message", {}).get("quoteToken", "")
            reply_msg = {"type": "text", "text": "賴翔德～我是不會屈服的"}
            if quote_token:
                reply_msg["quoteToken"] = quote_token
            reply_message(reply_token, [reply_msg])

        elif text == "張姉瑀":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/fish.jpg",
                "previewImageUrl": f"{BASE_URL}/fish.jpg"
            }])

        elif text == "林定緯":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/ding.jpg",
                "previewImageUrl": f"{BASE_URL}/ding.jpg"
            }])

        elif text == "午安猴":
            hao_images = ["hao.jpg", "hao2.jpg", "hao3.jpg"]
            img = random.choice(hao_images)
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/{img}",
                "previewImageUrl": f"{BASE_URL}/{img}"
            }])

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
                reply_message(reply_token, [{
                    "type": "text",
                    "text": "請在「小幫手」後面輸入你的問題！\n例如：小幫手 今天吃什麼好？"
                }])
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
                            {
                                "role": "system",
                                "content": "你是一個汽車業務團隊的AI小幫手，專門協助回答業務、銷售、客戶服務相關問題。請用繁體中文回答，回答簡潔有力、實用為主，平時還會有點搞笑帶給大家歡樂，也可以給予投資建議。"
                            },
                            {"role": "user", "content": question}
                        ]
                    }
                    resp = requests.post(groq_url, headers=groq_headers, json=groq_body)
                    answer = resp.json()["choices"][0]["message"]["content"]
                    reply_message(reply_token, [{"type": "text", "text": f"🤖 {answer}"}])
                except Exception as e:
                    reply_message(reply_token, [{"type": "text", "text": f"錯誤：{str(e)[:200]}"}])

        elif text == "推薦股票":
            stock_count[user_id] = stock_count.get(user_id, 0) + 1

            if stock_count[user_id] >= 3:
                stock_count[user_id] = 0
                reply_message(reply_token, [{"type": "text", "text": "戒斷當沖 是你唯一選擇"}])
            else:
                stocks = [
                    ("2330", "台積電"), ("2317", "鴻海"), ("2454", "聯發科"),
                    ("2382", "廣達"), ("2308", "台達電"), ("2303", "聯電"),
                    ("2881", "富邦金"), ("2882", "國泰金"), ("2886", "兆豐金"),
                    ("6669", "緯穎"), ("3017", "奇鋐"), ("2395", "研華"),
                ]
                code, name = random.choice(stocks)
                reply_message(reply_token, [{
                    "type": "text",
                    "text": f"📈 今日推薦股票\n\n【{code} {name}】\n\n⚠️ 僅供娛樂，不構成投資建議！"
                }])

    return "OK", 200


# ── 固定提醒路由 ──────────────────────────────────────
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


# ── 續保提醒觸發路由 ──────────────────────────────────
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


# ── 業績速報推播路由（每天 10:15 起每 15 分鐘觸發，成功後當天不再重複）────
_speed_report_pushed_date = None

@app.route("/push-speed-report", methods=["GET"])
def push_speed_report():
    global _speed_report_pushed_date
    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401

    from datetime import datetime, timezone, timedelta
    tw_now = datetime.now(timezone(timedelta(hours=8)))
    today_str = tw_now.strftime("%Y-%m-%d")

    # 今天已推播過，直接跳過
    if _speed_report_pushed_date == today_str:
        print(f"[SKIP] {today_str} 已推播過")
        return "Already pushed today", 200

    try:
        from drive_reader import get_speed_report, format_speed_report_message
        report = get_speed_report()

        # 確認速報日期在允許範圍內（今天或前一個工作日）
        report_date = report.get("date", "")
        report_date_str = f"{report_date[:4]}-{report_date[4:6]}-{report_date[6:]}" if len(report_date) == 8 else ""

        tw_today = tw_now.date()
        allowed = {(tw_today - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(3)}

        if report_date_str not in allowed:
            print(f"[WAIT] 速報日期 {report_date_str}，等待最新速報")
            return f"Waiting for latest report (got {report_date_str})", 200

        message = format_speed_report_message(report)
        resp = push_message(HSIN_USER_ID, [{"type": "text", "text": message}])

        if resp.status_code == 200:
            _speed_report_pushed_date = today_str
            print(f"[OK] 推播成功 {today_str}，速報日期 {report_date_str}")

        return f"OK: {resp.status_code}", 200

    except Exception as e:
        import traceback
        print(f"[ERROR] push_speed_report:\n{traceback.format_exc()}")
        return f"Error: {e}", 500


# ── Drive 測試路由 ────────────────────────────────────
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


# ── 基本路由 ──────────────────────────────────────────
@app.route("/", methods=["GET"])
def index():
    return "LINE Bot is running!", 200

@app.route("/img/<filename>")
def serve_image(filename):
    return send_from_directory(".", filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
from dashboard_routes import dashboard_bp
app.register_blueprint(dashboard_bp)
