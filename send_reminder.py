import json
import os
import sys
import requests
from datetime import datetime
import pytz

def send_message(token, group_id, message):
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = {
        "to": group_id,
        "messages": [{"type": "text", "text": message}]
    }
    response = requests.post(url, headers=headers, json=body)
    print(f"Status: {response.status_code}, Response: {response.text}")
    return response.status_code

def get_reminder_key():
    manual_key = os.environ.get("REMINDER_KEY", "")
    if manual_key:
        return manual_key

    tz = pytz.timezone("Asia/Taipei")
    now = datetime.now(tz)
    current_mins = now.hour * 60 + now.minute
    weekday = now.isoweekday()

    mttsat = {1, 2, 4, 6}
    schedules = [
        (8*60+45, mttsat, "morning_schedule"),
        (17*60+15, mttsat, "evening_schedule"),
        (12*60, set(range(1, 8)), "check_leads"),
        (9*60, {1}, "weekly_update"),
    ]

    for target_mins, days, key in schedules:
        if weekday in days and abs(current_mins - target_mins) <= 10:
            return key

    return None

def main():
    token = os.environ.get("LINE_TOKEN")
    if not token:
        print("Error: LINE_TOKEN not set")
        sys.exit(1)

    with open("reminders.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    reminder_key = get_reminder_key()
    if not reminder_key:
        print("No matching reminder for current time")
        sys.exit(0)

    reminders = config["reminders"]
    groups = config["groups"]

    if reminder_key not in reminders:
        print(f"Reminder key '{reminder_key}' not found")
        sys.exit(1)

    reminder = reminders[reminder_key]
    message = reminder["message"]
    target_groups = reminder.get("groups", list(groups.keys()))

    print(f"Sending reminder: {reminder_key}")
    for group_key in target_groups:
        if group_key in groups:
            status = send_message(token, groups[group_key], message)
            print(f"Sent to {group_key}: HTTP {status}")

if __name__ == "__main__":
    main()
