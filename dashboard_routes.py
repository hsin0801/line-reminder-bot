"""
歸仁 / 永康 / 法人 儀表板 - 路由模組
把這三個 blueprint 都掛進你現有的 app.py：

    from dashboard_routes import dashboard_bp, yongkang_bp, faren_bp
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(yongkang_bp)
    app.register_blueprint(faren_bp)

網址：
    https://你的render網址/gueiren-dashboard/     歸仁
    https://你的render網址/yongkang-dashboard/    永康
    https://你的render網址/faren-dashboard/       法人（永康+歸仁合併）

需要另外設定的 cron-job.org 排程，各自都要各打一次refresh：
    GET https://你的render網址/gueiren-dashboard/refresh?token=你設定的密鑰
    GET https://你的render網址/yongkang-dashboard/refresh?token=你設定的密鑰
    GET https://你的render網址/faren-dashboard/refresh?token=你設定的密鑰
（法人的refresh內部會各自重新呼叫歸仁跟永康的解析，所以三個都排程也不會互相打架，
 但如果想省事，法人排在歸仁、永康都跑完之後幾分鐘再排也可以）
"""

import os
import json
from flask import Blueprint, render_template, jsonify, request

from dashboard_parser import build_dashboard_data, DATA_FILE
from yongkang_parser import build_yongkang_data, DATA_FILE as YONGKANG_DATA_FILE
from faren_parser import build_faren_data, DATA_FILE as FAREN_DATA_FILE

REFRESH_TOKEN = os.environ.get("DASHBOARD_REFRESH_TOKEN", "")


def _load_cached(path):
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _check_token():
    token = request.args.get("token", "")
    return not (REFRESH_TOKEN and token != REFRESH_TOKEN)


# ===== 歸仁 =====
dashboard_bp = Blueprint(
    "gueiren_dashboard", __name__, template_folder="templates", url_prefix="/gueiren-dashboard",
)


@dashboard_bp.route("/")
def show_dashboard():
    data = _load_cached(DATA_FILE)
    if data is None:
        return "資料還沒有產生，請先呼叫 /gueiren-dashboard/refresh 一次", 503
    return render_template("dashboard.html", data=data)


@dashboard_bp.route("/refresh")
def refresh():
    if not _check_token():
        return "unauthorized", 401
    data = build_dashboard_data()
    return jsonify({"status": "ok", "updated_at": data["updated_at"], "source_file": data["source_file"]})


@dashboard_bp.route("/data.json")
def raw_data():
    data = _load_cached(DATA_FILE)
    if data is None:
        return jsonify({"error": "no data yet"}), 503
    return jsonify(data)


# ===== 永康 =====
yongkang_bp = Blueprint(
    "yongkang_dashboard", __name__, template_folder="templates", url_prefix="/yongkang-dashboard",
)


@yongkang_bp.route("/")
def show_yongkang():
    data = _load_cached(YONGKANG_DATA_FILE)
    if data is None:
        return "資料還沒有產生，請先呼叫 /yongkang-dashboard/refresh 一次", 503
    return render_template("yongkang_dashboard.html", data=data)


@yongkang_bp.route("/refresh")
def refresh_yongkang():
    if not _check_token():
        return "unauthorized", 401
    data = build_yongkang_data()
    return jsonify({"status": "ok", "updated_at": data["updated_at"], "source_file": data["source_file"]})


@yongkang_bp.route("/data.json")
def raw_data_yongkang():
    data = _load_cached(YONGKANG_DATA_FILE)
    if data is None:
        return jsonify({"error": "no data yet"}), 503
    return jsonify(data)


# ===== 法人 =====
faren_bp = Blueprint(
    "faren_dashboard", __name__, template_folder="templates", url_prefix="/faren-dashboard",
)


@faren_bp.route("/")
def show_faren():
    data = _load_cached(FAREN_DATA_FILE)
    if data is None:
        return "資料還沒有產生，請先呼叫 /faren-dashboard/refresh 一次", 503
    return render_template("faren_dashboard.html", data=data)


@faren_bp.route("/refresh")
def refresh_faren():
    if not _check_token():
        return "unauthorized", 401
    data = build_faren_data()
    return jsonify({"status": "ok", "updated_at": data["updated_at"]})


@faren_bp.route("/data.json")
def raw_data_faren():
    data = _load_cached(FAREN_DATA_FILE)
    if data is None:
        return jsonify({"error": "no data yet"}), 503
    return jsonify(data)
