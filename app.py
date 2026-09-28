import os
import json
import requests
import random
import time
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

# 記憶體計數器（處理群組內部的次數計數）
stock_count = {}
po_count = {}


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
    if not LINE_TOKEN:
        print("[ERROR] LINE_TOKEN 環境變數未設定！")
        return None
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    body = {"to": to, "messages": messages}
    try:
        resp = requests.post(url, headers=headers, json=body, timeout=10)
        print(f"[LINE API] status={resp.status_code} body={resp.text[:300]}")
        return resp
    except Exception as e:
        import traceback
        print(f"[ERROR] push_message exception:\n{traceback.format_exc()}")
        return None

def quota_push(to, messages, label, priority="normal", size=None):
    """先問額度守門員，夠才送；送成功才記帳。回傳是否送出。"""
    import line_quota
    group_id = os.environ.get("REMINDER_GROUP_ID")
    ok, info = line_quota.check(group_id, pushes=1, priority=priority, size=size)
    if not ok:
        print(f"[SKIP] {label} 額度不足未發送（已用 {info['used']}/{info['limit']}）")
        return False
    resp = push_message(to, messages)
    if resp is None or resp.status_code != 200:
        return False
    line_quota.commit(info, pushes_sent=1, label=label)
    return True


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
                "text": "📊 配件資料：\nhttps://docs.google.com/spreadsheets/d/1p0XQBPX8B0fjMGHg40H-zdGIfvqDqUJT/edit?usp=sharing&ouid=109189035277985438460&rtpof=true&sd=true"
            }])

        elif text == "組合價":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/combination.jpg{cache_buster}",
                "previewImageUrl": f"{BASE_URL}/combination.jpg{cache_buster}"
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
                    import assistant
                    answer = assistant.answer(question)
                    reply_message(reply_token, [{"type": "text", "text": f"🤖 {answer}"}])
                except Exception as e:
                    print(f"[ASSISTANT] {e}")
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
    if request.args.get("secret", "") != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401
    default_config = {
        "groups": {"歸仁包廂": "Cac09c73b2a7562516bbd7516a9352a56"},
        "reminders": {
            "weekly_update": {"message": "📊 更新週邊指標及續保", "groups": ["歸仁包廂"]},
            "llc_reminder": {"message": "📋 LLC 今天記得完成！", "groups": ["歸仁包廂"]},
            "sunday_prospects": {
                "message": "📋 晚上盤點下週有望成交客戶\n@林定緯 @陳星佑",
                "groups": ["歸仁包廂"],
                "mentions": ["林定緯", "陳星佑"]
            }
        }
    }
    try:
        with open("reminders.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        if "reminders" not in config or "groups" not in config:
            raise ValueError("reminders.json 結構不完整")
    except Exception as e:
        print(f"[WARN] reminders.json 讀取失敗，自動還原預設值: {e}")
        config = default_config
        try:
            with open("reminders.json", "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as we:
            print(f"[ERROR] 無法寫入 reminders.json: {we}")

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
            quota_push(groups[group_key], [msg], f"提醒 {key}")

    return "OK", 200


# ── 8. 基本路由 ──────────────────────────────────────────
@app.route("/assistant-test", methods=["GET"])
def assistant_test():
    if request.args.get("secret", "") != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401
    import assistant
    q = request.args.get("q", "")
    if request.args.get("context") == "1":
        return assistant.get_context(), 200, {"Content-Type": "text/plain; charset=utf-8"}
    try:
        text = assistant.answer(q)
    except Exception as e:
        text = f"ERROR: {e}"
    return text, 200, {"Content-Type": "text/plain; charset=utf-8"}

@app.route("/quota-status", methods=["GET"])
def quota_status():
    if request.args.get("secret", "") != os.environ.get("CRON_SECRET", ""):
        return "Unauthorized", 401
    import line_quota
    info = line_quota.status(os.environ.get("REMINDER_GROUP_ID"))
    return json.dumps(info, ensure_ascii=False, indent=2), 200, {"Content-Type": "application/json; charset=utf-8"}

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
