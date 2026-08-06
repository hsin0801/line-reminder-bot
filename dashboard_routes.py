"""
歸仁 / 永康 / 法人 儀表板 - 路由模組
把這四個 blueprint 都掛進你現有的 app.py：

    from dashboard_routes import dashboard_bp, yongkang_bp, faren_bp, combined_bp
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(yongkang_bp)
    app.register_blueprint(faren_bp)
    app.register_blueprint(combined_bp)

網址：
    https://你的render網址/dashboard/              整合頁面（單一網址，分頁籤切換歸仁/永康/法人，推薦用這個）
    https://你的render網址/gueiren-dashboard/     歸仁（獨立網址，仍保留）
    https://你的render網址/yongkang-dashboard/    永康（獨立網址，仍保留）
    https://你的render網址/faren-dashboard/       法人（獨立網址，仍保留）

需要另外設定的 cron-job.org 排程，最簡單的做法是只排一個：
    GET https://你的render網址/dashboard/refresh?token=你設定的密鑰
（這個會依序刷新歸仁→永康→法人三份資料，一次搞定，不用排三個）
"""

import os
import json
from flask import Blueprint, render_template, jsonify, request

from dashboard_parser import build_dashboard_data, DATA_FILE, backfill_full_history as gueiren_backfill_full_history, reset_order_history as gueiren_reset_order_history
from yongkang_parser import build_yongkang_data, DATA_FILE as YONGKANG_DATA_FILE, backfill_full_history as yongkang_backfill_full_history, reset_order_history as yongkang_reset_order_history
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
        data = build_dashboard_data()  # 資料不見了(例如Render休眠後清空)，自動重新產生
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


@dashboard_bp.route("/backfill")
def backfill():
    """完整回溯歷史(CR-V PET/e:HEV當日值+其他車型)。因為檔案數量多(140+份)，
    可能要重複呼叫這個網址好幾次才能跑完，每次呼叫回傳的 done:false 代表還沒跑完，
    要繼續打同一個網址；done:true 代表全部處理完成。"""
    if not _check_token():
        return "unauthorized", 401
    result = gueiren_backfill_full_history()
    return jsonify(result)


@dashboard_bp.route("/reset-history")
def reset_history():
    """清空Drive上的每日訂單快照歷史，重新開始回溯用。修正CR-V欄位驗證邏輯後，
    之前回溯進去的1~4月錯誤資料需要先清掉，再重新呼叫 /backfill 才會是乾淨的資料。"""
    if not _check_token():
        return "unauthorized", 401
    result = gueiren_reset_order_history()
    return jsonify(result)


@dashboard_bp.route("/debug-history")
def debug_history():
    """除錯用：直接看Drive上存的原始每日快照歷史，某個人在指定日期範圍的資料。
    用法：/gueiren-dashboard/debug-history?name=張姉瑀&from=2026-07-10&to=2026-07-19"""
    from dashboard_parser import load_order_history
    name = request.args.get("name", "")
    date_from = request.args.get("from", "0000-00-00")
    date_to = request.args.get("to", "9999-99-99")
    history = load_order_history()
    result = {}
    for d in sorted(history.keys()):
        if date_from <= d <= date_to:
            person_data = history[d].get(name)
            if person_data is not None:
                result[d] = person_data
    return jsonify({"name": name, "total_dates_in_history": len(history), "matched": result})


# ===== 永康 =====
yongkang_bp = Blueprint(
    "yongkang_dashboard", __name__, template_folder="templates", url_prefix="/yongkang-dashboard",
)


@yongkang_bp.route("/")
def show_yongkang():
    data = _load_cached(YONGKANG_DATA_FILE)
    if data is None:
        data = build_yongkang_data()
    return render_template("yongkang_dashboard.html", data=data)


@yongkang_bp.route("/refresh")
def refresh_yongkang():
    if not _check_token():
        return "unauthorized", 401
    data = build_yongkang_data()
    return jsonify({"status": "ok", "updated_at": data["updated_at"], "source_file": data["source_file"]})


@yongkang_bp.route("/backfill")
def backfill_yongkang():
    if not _check_token():
        return "unauthorized", 401
    result = yongkang_backfill_full_history()
    return jsonify(result)


@yongkang_bp.route("/reset-history")
def reset_history_yongkang():
    if not _check_token():
        return "unauthorized", 401
    result = yongkang_reset_order_history()
    return jsonify(result)


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
        data = build_faren_data()  # build_faren_data內部會各自呼叫歸仁/永康的build，資料不見時一併補齊
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


# ===== 整合頁面（單一網址，用分頁籤切換三個視圖）=====
combined_bp = Blueprint(
    "combined_dashboard", __name__, template_folder="templates", url_prefix="/dashboard",
)


@combined_bp.route("/")
def show_combined():
    g_data = _load_cached(DATA_FILE)
    y_data = _load_cached(YONGKANG_DATA_FILE)
    f_data = _load_cached(FAREN_DATA_FILE)
    if g_data is None or y_data is None or f_data is None:
        # 資料不見了(例如Render休眠後本機檔案被清空)，自動重新產生。
        # 注意：只呼叫 build_faren_data() 一次就好——它內部本來就會重新解析
        # 歸仁跟永康並各自存檔，如果這裡再額外呼叫 build_dashboard_data()/
        # build_yongkang_data()，等於同一份資料在同一個請求裡被解析兩次，
        # 記憶體用量直接翻倍，先前就是這樣被 Render 判定OOM砍掉的。
        f_data = build_faren_data()
        g_data = _load_cached(DATA_FILE)      # build_faren_data() 內部已經順便寫好這份快取了
        y_data = _load_cached(YONGKANG_DATA_FILE)
    return render_template("combined_dashboard.html", g=g_data, y=y_data, f=f_data)


@combined_bp.route("/refresh")
def refresh_combined():
    """一次刷新三份資料（歸仁→永康→法人，法人最後跑因為需要前兩者的結果）。"""
    if not _check_token():
        return "unauthorized", 401
    g_data = build_dashboard_data()
    y_data = build_yongkang_data()
    f_data = build_faren_data()
    return jsonify({
        "status": "ok",
        "gueiren_updated_at": g_data["updated_at"],
        "yongkang_updated_at": y_data["updated_at"],
        "faren_updated_at": f_data["updated_at"],
    })

# ===== 戰情室風格儀表板（動態版，從 API 拉資料）=====
# ── 歸仁績效指標 API ──
from flask import Blueprint as _BP
_guiren_kpi_bp = _BP("guiren_kpi", __name__, url_prefix="/guiren-kpi")
_guiren_kpi_cache = {"data": None, "ts": 0}

@_guiren_kpi_bp.route("/data.json")
def guiren_kpi_data():
    import time, logging
    now = time.time()
    if _guiren_kpi_cache["data"] and now - _guiren_kpi_cache["ts"] < 300:
        return jsonify(_guiren_kpi_cache["data"])
    try:
        from guiren_kpi_reader import get_guiren_kpi
        from dashboard_parser import get_drive_service
        data = get_guiren_kpi(get_drive_service())
        _guiren_kpi_cache["data"] = data
        _guiren_kpi_cache["ts"] = now
        return jsonify(data)
    except Exception as e:
        logging.error(f"[guiren_kpi] {e}")
        if _guiren_kpi_cache["data"]:
            return jsonify(_guiren_kpi_cache["data"])
        return jsonify({"error": str(e)}), 503

warroom_bp = Blueprint(
    "warroom", __name__, template_folder="templates", url_prefix="/warroom",
)


@warroom_bp.route("/")
def show_warroom():
    """戰情室風格儀表板，前端 JS 自動呼叫各 /data.json 取得資料，不需要 server-side 渲染。"""
    return render_template("warroom_dashboard.html")
