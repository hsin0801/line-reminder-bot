import os
import json
import requests
import random
from datetime import datetime
from flask import Flask, request, send_from_directory

app = Flask(__name__)

LINE_TOKEN = os.environ.get("LINE_TOKEN")
BASE_URL = "https://line-reminder-bot-gj9p.onrender.com/img"
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

        if text == "內促":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/may.jpg",
                "previewImageUrl": f"{BASE_URL}/may.jpg"
            }])

        elif text == "SP":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📄 SP 活動資料：\nhttps://drive.google.com/file/d/1yAwAcuvOseGMe5fubRTliHsH87K2avmP/view?usp=sharing"
            }])

        elif text == "配件":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📊 配件資料：\nhttps://docs.google.com/spreadsheets/d/179_z6rI0ZpyQHGfjQ_3qQUZAufQXO0kC/edit?usp=sharing"
            }])

        elif text == "組合價":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": f"{BASE_URL}/combination.jpg",
                "previewImageUrl": f"{BASE_URL}/combination.jpg"
            }])

        elif text == "紀柏州":
            user_id = event.get("source", {}).get("userId", "unknown")
            group_id = event.get("source", {}).get("groupId", "")
            quote_token = event.get("message", {}).get("quoteToken", "")
            po_count[user_id] = po_count.get(user_id, 0) + 1

            if po_count[user_id] >= 3:
                po_count[user_id] = 0

                display_name = "你"
                if group_id and user_id != "unknown":
                    try:
                        profile_url = f"https://api.line.me/v2/bot/group/{group_id}/member/{user_id}"
                        headers = {"Authorization": f"Bearer {LINE_TOKEN}"}
                        resp = requests.get(profile_url, headers=headers)
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
                        "mentionees": [
                            {
                                "index": 0,
                                "length": len(mention_text),
                                "type": "user",
                                "userId": user_id
                            }
                        ]
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
            reply_msg = {
                "type": "text",
                "text": "賴翔德～我是不會屈服的"
            }
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

        elif text == "接龍":
            reply_message(reply_token, [{
                "type": "text",
                "text": "宗鑫\n定緯\n適緯\n珈微\n建道\n星佑\n姉瑀\n文智\n明憬"
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
                        "model": "llama-3.3-70b-versatile",
                        "messages": [
                            {
                                "role": "system",
                                "content": "你是一個財經專家，給予投資建議。"
                            },
                            {"role": "user", "content": question}
                        ]
                    }
                    resp = requests.post(groq_url, headers=groq_headers, json=groq_body)
                    answer = resp.json()["choices"][0]["message"]["content"]
                    reply_message(reply_token, [{
                        "type": "text",
                        "text": f"🤖 {answer}"
                    }])
                except Exception as e:
                    reply_message(reply_token, [{
                        "type": "text",
                        "text": f"錯誤：{str(e)[:200]}"
                    }])

        elif text == "推薦股票":
            user_id = event.get("source", {}).get("userId", "unknown")
            stock_count[user_id] = stock_count.get(user_id, 0) + 1

            if stock_count[user_id] >= 3:
                stock_count[user_id] = 0
                reply_message(reply_token, [{
                    "type": "text",
                    "text": "戒斷當沖 是你唯一選擇"
                }])
            else:
                stocks = [
                    ("2330", "台積電"), ("2317", "鴻海"), ("2454", "聯發科"),
                    ("2382", "廣達"), ("2308", "台達電"), ("2303", "聯電"),
                    ("2881", "富邦金"), ("2882", "國泰金"), ("2886", "兆豐金"),
                    ("1301", "台塑"), ("1303", "南亞"), ("2002", "中鋼"),
                    ("2412", "中華電"), ("3008", "大立光"), ("2891", "中信金"),
                    ("2884", "玉山金"), ("2885", "元大金"), ("5871", "中租-KY"),
                    ("2395", "研華"), ("3711", "日月光投控"), ("2379", "瑞昱"),
                    ("3045", "台灣大"), ("4904", "遠傳"), ("2357", "華碩"),
                    ("2353", "宏碁"), ("2207", "和泰車"), ("1216", "統一"),
                    ("2912", "統一超"), ("2887", "台新金"), ("2890", "永豐金"),
                    ("2892", "第一金"), ("5880", "合庫金"), ("2883", "開發金"),
                    ("2880", "華南金"), ("2888", "新光金"), ("6505", "台塑化"),
                    ("1326", "台化"), ("2347", "聯強"), ("2356", "英業達"),
                    ("2376", "技嘉"), ("2377", "微星"), ("2385", "群光"),
                    ("2409", "友達"), ("3481", "群創"), ("2449", "京元電子"),
                    ("2458", "義隆"), ("2474", "可成"), ("2492", "華新科"),
                    ("2633", "台灣高鐵"), ("2801", "彰銀"), ("3034", "聯詠"),
                    ("2327", "國巨"), ("2408", "南亞科"), ("5269", "祥碩"),
                    ("6415", "矽力-KY"), ("6239", "力成"), ("3037", "欣興"),
                    ("4958", "臻鼎-KY"), ("2383", "台光電"), ("3044", "健鼎"),
                    ("6271", "同欣電"), ("2345", "智邦"), ("4938", "和碩"),
                    ("2354", "鴻準"), ("2352", "佳世達"), ("3017", "奇鋐"),
                    ("6669", "緯穎"), ("3687", "貿聯-KY"), ("2360", "致茂"),
                    ("3019", "亞光"), ("2368", "金像電"), ("3533", "嘉澤"),
                    ("6176", "瑞儀"), ("6446", "藥華藥"), ("9910", "豐泰"),
                    ("2105", "正新"), ("1402", "遠東新"), ("2603", "長榮"),
                    ("2609", "陽明"), ("2615", "萬海"), ("2610", "華航"),
                    ("2618", "長榮航"), ("2201", "裕隆"), ("2204", "中華汽車"),
                    ("9914", "美利達"), ("9921", "巨大"), ("1590", "亞德客-KY"),
                    ("2049", "上銀"), ("3231", "緯創"), ("2498", "宏達電"),
                    ("1217", "愛之味"), ("1210", "大成"), ("1229", "聯華"),
                    ("1232", "大統益"), ("1234", "黑松"), ("2548", "華固"),
                    ("2542", "興富發"), ("5534", "長虹"), ("2915", "潤泰全"),
                    ("2903", "遠百"), ("3443", "創意"), ("6409", "旭隼"),
                    ("6533", "晶心科"), ("2606", "裕民"), ("2823", "中壽"),
                    ("2867", "三商壽"), ("2834", "臺企銀"), ("6116", "彩晶"),
                    ("1533", "車王電"), ("2227", "裕日車"), ("1305", "華夏"),
                    ("1312", "國喬"), ("1314", "中石化"), ("2404", "漢唐"),
                    ("4912", "聯陽"), ("6278", "台表科"), ("3673", "TPK宸鴻"),
                ]
                code, name = random.choice(stocks)
                reply_message(reply_token, [{
                    "type": "text",
                    "text": f"📈 今日推薦股票\n\n【{code} {name}】\n\n⚠️ 僅供娛樂，不構成投資建議！"
                }])

    return "OK", 200

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

@app.route("/", methods=["GET"])
def index():
    return "LINE Bot is running!", 200

@app.route("/img/<filename>")
def serve_image(filename):
    return send_from_directory(".", filename)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
