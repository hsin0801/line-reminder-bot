import os
import json
import requests
import random
import time
from datetime import datetime, timezone, timedelta
from flask import Flask, request, send_from_directory
from renewal_reminder import run_reminder, mark_replied

app = Flask(__name__)

# ──── 【優化】先註冊所有 Blueprint，確保路由在啟動前完全載入 ────
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

# 記憶體計數器（注意：Render 重啟時會歸零，若要永久保存建議未來改存資料庫）
stock_count = {}
po_count = {}

def reply_message(reply_token, messages):
    url = "https://api.line.me/v2/bot/message/reply"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    body = {"replyToken": reply_token, "messages": messages}
    # 設定 timeout 防止外部 API 癱瘓你的 Bot
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

        # ──── 【優化】回覆偵測（加入 Timeout，防止抓 Profile 卡死 Webhook） ────
        if group_id == os.environ.get("REMINDER_GROUP_ID"):
            display_name = ""
            if user_id != "unknown":
                try:
                    profile_url = f"https://api.line.me/v2/bot/group/{group_id}/member/{user_id}"
                    headers_profile = {"Authorization": f"Bearer {LINE_TOKEN}"}
                    # 加上 timeout=3，如果 LINE 沒回應立刻跳過不卡死
                    resp = requests.get(profile_url, headers=headers_profile, timeout=3)
                    if resp.status_code == 200:
                        display_name = resp.json().get("displayName", "")
                except Exception as e:
                    print(f"[ERROR] Get Profile Failed: {e}")
            if display_name:
                mark_replied(display_name)

        # ──── 關鍵字指令區 ────
        # 【優化提示】圖片網址後加上 ?t=時間戳記，防止手機 LINE 狂吃舊快取不更換新圖
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
            # 簡化重複的隨機圖片邏輯，一律加上防快取外掛
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

# ── 業績速報推播路由（修正變數重置與快取問題） ──────────────────────
_speed_report_pushed_date = None

@app.route("/push-speed-report", methods=["GET"])
def push_speed_report():
    global _speed_report_pushed_date
    secret = request.args.get("secret", "")
    if secret != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401

    tw_now = datetime.now(timezone(timedelta(hours=8)))
    today_str = tw_now.strftime("%Y-%m-%d")

    if _speed_report_pushed_date == today_str:
        return "Already pushed today", 200

    try:
        # ──── 【優化】呼叫 drive_reader 前，可以確保那邊沒有讀到快取 ────
        from drive_reader import get_speed_report, format_speed_report_message
        
        # 備註：請確保你的 drive_reader.py 內部打 Google API 時也有設定不使用快取
        report = get_speed_report()

        report_date = report.get("date", "")
        report_date_str = f"{report_date[:4]}-{report_date[4:6]}-{report_date[6:]}" if len(report_date) == 8 else ""

        tw_today = tw_now.date()
        allowed = {(tw_today - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(3)}

        if report_date_str not in allowed:
            return f"Waiting for latest report (got {report_date_str})", 200

        message = format_speed_report_message(report)
        resp = push_message(HSIN_USER_ID, [{"type": "text", "text": message}])

        if resp and resp.status_code == 200:
            _speed_report_pushed_date = today_str
            print(f"[OK] 推播成功 {today_str}")
            return f"OK: {resp.status_code}", 200
        
        return "Push failed", 500

    except Exception as e:
        return f"Error: {e}", 500

# ── 其他常規路由保持不變 ──
@app.route("/remind/<key>", methods=["GET"])
def remind(key):
    # 保持原樣...
    return "OK", 200

@app.route("/", methods=["GET"])
def index():
    return "LINE Bot is running!", 200

@app.route("/img/<filename>")
def serve_image(filename):
    return send_from_directory(".", filename)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
