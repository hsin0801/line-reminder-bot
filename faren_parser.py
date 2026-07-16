"""
法人儀表板 - 歸仁 + 永康 合併
分別呼叫兩邊各自的 build_*_data()，然後做加總。
不重新抓解析邏輯，直接復用兩邊已經算好的結果，降低出錯風險
（歸仁跟永康的解析程式已經各自獨立驗證過）。
"""

import json
from datetime import datetime

import dashboard_parser as gueiren
import yongkang_parser as yongkang

DATA_FILE = "faren_dashboard_data.json"


def build_faren_data():
    g = gueiren.build_dashboard_data()
    y = yongkang.build_yongkang_data()

    # ---- 總計 ----
    team_total = g["team_total_ytd_registration"] + y["team_total_ytd_registration"]

    # ---- 各據點小計（本身就有現成的team_total）----
    branch_totals = {
        "歸仁": g["team_total_ytd_registration"],
        "永康": y["team_total_ytd_registration"],
    }

    # ---- 各業務員領牌台數，合併兩據點（正常不會撞名，如果真的撞名會加在一起）----
    item1_merged = {}
    for name, val in g["item1_ytd_registration"].items():
        item1_merged[("歸仁", name)] = val
    for name, val in y["item1_ytd_registration"].items():
        item1_merged[("永康", name)] = val

    # ---- 各車型合併銷售台數（兩據點加總，不分業務員，看車型層級的公司總表現）----
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

    # ---- 本月訂單/領牌進度：兩據點加總 ----
    def sum_month_progress(mp):
        订单 = sum(v.get('訂單', 0) for v in mp.values())
        领牌 = sum(v.get('領牌', 0) for v in mp.values())
        return {'訂單': 订单, '領牌': 领牌}

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
        "gueiren_dept_totals": g["item1_dept_totals"],
        "yongkang_dept_totals": y["item1_dept_totals"],
        "gueiren_item1": g["item1_ytd_registration"],
        "yongkang_item1": y["item1_ytd_registration"],
        "gueiren_item4": g["item4_ytd_by_model"],
        "yongkang_item4": y["item4_ytd_by_model"],
        "model_totals_company": model_totals_company,
        "month_progress_company": month_progress_company,
        "gueiren_month_progress": g["month_progress"],
        "yongkang_month_progress": y["month_progress"],
        "gueiren_team_structure": g["team_structure"],
        "yongkang_team_structure": y["team_structure"],
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return data


if __name__ == "__main__":
    result = build_faren_data()
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
