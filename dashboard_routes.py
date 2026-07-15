"""
歸仁儀表板 - 路由模組
把這個 blueprint 掛進你現有的 app.py：

    from dashboard_routes import dashboard_bp
    app.register_blueprint(dashboard_bp)

之後網址就是：https://你的render網址/gueiren-dashboard
（路徑名稱可以自己改，避免跟現有 LINE Bot 的路由撞名）

需要另外設定的 cron-job.org 排程（跟你現有的做法一樣）：
    每天早上（建議業績日報表寄到之後，例如10:30）打一次：
    GET https://你的render網址/gueiren-dashboard/refresh?token=你設定的密鑰

用 token 是為了避免這個刷新網址被別人亂打導致重複下載。
"""

import os
import json
from flask import Blueprint, render_template, jsonify, request

from dashboard_parser import build_dashboard_data, DATA_FILE

dashboard_bp = Blueprint(
    "gueiren_dashboard",
    __name__,
    template_folder="templates",
    url_prefix="/gueiren-dashboard",
)

REFRESH_TOKEN = os.environ.get("DASHBOARD_REFRESH_TOKEN", "")


def load_cached_data():
    if not os.path.exists(DATA_FILE):
        return None
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


@dashboard_bp.route("/")
def show_dashboard():
    data = load_cached_data()
    if data is None:
        return "資料還沒有產生，請先呼叫 /gueiren-dashboard/refresh 一次", 503
    return render_template("dashboard.html", data=data)


@dashboard_bp.route("/refresh")
def refresh():
    token = request.args.get("token", "")
    if REFRESH_TOKEN and token != REFRESH_TOKEN:
        return "unauthorized", 401
    data = build_dashboard_data()
    return jsonify({"status": "ok", "updated_at": data["updated_at"], "source_file": data["source_file"]})


@dashboard_bp.route("/data.json")
def raw_data():
    data = load_cached_data()
    if data is None:
        return jsonify({"error": "no data yet"}), 503
    return jsonify(data)
