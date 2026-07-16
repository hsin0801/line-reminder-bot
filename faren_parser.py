"""
法人儀表板 - 歸仁 + 永康 合併
分別呼叫兩邊各自的 build_*_data()，然後做加總。
不重新抓解析邏輯，直接復用兩邊已經算好的結果，降低出錯風險
（歸仁跟永康的解析程式已經各自獨立驗證過）。

法人頁面呈現邏輯：不分「歸仁區塊/永康區塊」，而是把兩邊的課別當成同一層級的
五個課（歸仁一課、歸仁二課、永康一課、永康二課、永康三課），所有業務員合併
排序/分類顯示，要看據點各自的細節就直接去歸仁/永康各自的分頁。
"""

import json
from datetime import datetime

import dashboard_parser as gueiren
import yongkang_parser as yongkang

DATA_FILE = "faren_dashboard_data.json"


def build_faren_data():
    g = gueiren.build_dashboard_data()
    y = yongkang.build_yongkang_data()

    # ---- 統一的五課結構（歸仁一課、歸仁二課、永康一課、永康二課、永康三課）----
    # 注意：兩邊的 team_structure key 本身就已經是「歸仁一課」「永康一課」這種完整名稱，
    # 不需要再加前綴（加了會變成「歸仁歸仁一課」的重複bug）。
    faren_team_structure = {}
    faren_team_structure.update(g["team_structure"])
    faren_team_structure.update(y["team_structure"])

    # ---- 總計 ----
    team_total = g["team_total_ytd_registration"] + y["team_total_ytd_registration"]
    branch_totals = {"歸仁": g["team_total_ytd_registration"], "永康": y["team_total_ytd_registration"]}

    # ---- 五課的領牌小計 ----
    dept_totals = {}
    dept_totals.update(g["item1_dept_totals"])
    dept_totals.update(y["item1_dept_totals"])

    # ---- 全部業務員合併（領牌台數），正常不會撞名----
    item1_all = {}
    item1_all.update(g["item1_ytd_registration"])
    item1_all.update(y["item1_ytd_registration"])

    # ---- 各業務員/各車型銷售台數，合併 ----
    item4_all = {}
    item4_all.update(g["item4_ytd_by_model"])
    item4_all.update(y["item4_ytd_by_model"])

    # ---- 本月訂單/領牌進度，合併 ----
    month_progress_all = {}
    month_progress_all.update(g["month_progress"])
    month_progress_all.update(y["month_progress"])

    # ---- 最後一筆訂單追蹤，合併 ----
    last_order_tracking_all = {}
    last_order_tracking_all.update(g["last_order_tracking"])
    last_order_tracking_all.update(y["last_order_tracking"])

    # ---- 各車型合併銷售台數（公司整體，不分業務員）----
    def merge_model_totals(item4_gueiren, item4_yongkang):
        totals = {}
        for models in item4_gueiren.values():
            for model, v in models.items():
                totals[model] = totals.get(model, 0) + v
        for models in item4_yongkang.values():
            for model, v in models.items():
                totals[model] = totals.get(model, 0) + v
        return totals

    model_totals_company = merge_model_totals(g["item4_ytd_by_model"], y["item4_ytd_by_model"])

    # ---- 本月訂單/領牌進度：據點對比（公司整體視角用）----
    def sum_month_progress(mp):
        return {'訂單': sum(v.get('訂單', 0) for v in mp.values()),
                '領牌': sum(v.get('領牌', 0) for v in mp.values())}

    month_progress_company = {
        "歸仁": sum_month_progress(g["month_progress"]),
        "永康": sum_month_progress(y["month_progress"]),
    }

    data = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "gueiren_source_file": g["source_file"],
        "yongkang_source_file": y["source_file"],
        "team_total_ytd_registration": team_total,
        "branch_totals": branch_totals,
        "faren_team_structure": faren_team_structure,
        "dept_totals": dept_totals,
        "item1_all": item1_all,
        "item4_all": item4_all,
        "month_progress_all": month_progress_all,
        "last_order_tracking_all": last_order_tracking_all,
        "model_totals_company": model_totals_company,
        "month_progress_company": month_progress_company,
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return data


if __name__ == "__main__":
    result = build_faren_data()
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
