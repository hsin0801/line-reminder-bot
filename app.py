import os
import json
import requests
from flask import Flask, request

app = Flask(__name__)

LINE_TOKEN = os.environ.get("LINE_TOKEN")
NEICU_IMAGE_URL = "https://raw.githubusercontent.com/hsin0801/line-reminder-bot/main/may.jpg"
SP_PDF_URL = "https://drive.google.com/file/d/1yAwAcuvOseGMe5fubRTliHsH87K2avmP/view?usp=sharing"

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
                "originalContentUrl": NEICU_IMAGE_URL,
                "previewImageUrl": NEICU_IMAGE_URL
            }])
        elif text == "SP":
            reply_message(reply_token, [{
                "type": "text",
                "text": f"📄 SP 活動資料：\n{SP_PDF_URL}"
            }])
        elif text == "配件":
            reply_message(reply_token, [{
                "type": "text",
                "text": "📊 配件資料：\nhttps://docs.google.com/spreadsheets/d/179_z6rI0ZpyQHGfjQ_3qQUZAufQXO0kC/edit?usp=sharing"
            }])
        elif text == "組合價":
            reply_message(reply_token, [{
                "type": "image",
                "originalContentUrl": "https://raw.githubusercontent.com/hsin0801/line-reminder-bot/main/combination.jpg",
"previewImageUrl": "https://raw.githubusercontent.com/hsin0801/line-reminder-bot/main/combination.jpg"
            }])  
            
    return "OK", 200

@app.route("/", methods=["GET"])
def index():
    return "LINE Bot is running!", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
