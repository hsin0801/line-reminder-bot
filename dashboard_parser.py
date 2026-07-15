"""
歸仁儀表板 - 資料解析模組（v2，沿用現有 drive_reader.py 的 Drive 連線邏輯）
不需要新的環境變數，直接複用 download_file()。
"""

import io
import os
import re
import json
from datetime import date, datetime
from collections import defaultdict

import openpyxl

from drive_reader import download_file, DAILY_REPORT_FOLDER_ID

DATA_FILE = "dashboard_data.json"


def _parse_filename_date(filename):
    """從檔名解析出這份報表對應的最後日期，回傳 (month, day) 供排序比較。
    處理常見格式："115 01 02"、"115  07 15"、"115 0122"、"115 07 11-13"、
    "115 0124-26"、"115 0227-0302"（跨月區間，取結束日期）等。
    解析失敗回傳 (0, 0)（排最後面，不會被誤判為最新）。
    """
    # 找到「115」之後、副檔名之前的部分
    m = re.search(r'115\s*([0-9\-\s]+)', filename)
    if not m:
        return (0, 0)
    raw = m.group(1)
    raw = raw.replace(' ', '').strip('-')
    if not raw:
        return (0, 0)

    parts = raw.split('-')
    try:
        if len(parts) == 1:
            digits = parts[0]
            if len(digits) < 3:
                return (0, 0)
            # 補齊成 MMDD（3碼視為 M+DD，4碼視為 MMDD）
            if len(digits) == 3:
                mm, dd = digits[0], digits[1:]
            else:
                mm, dd = digits[:2], digits[2:4]
            return (int(mm), int(dd))
        else:
            # 區間取「結束日期」那一段
            end = parts[-1]
            start = parts[0]
            if len(end) >= 4:
                mm, dd = end[:2], end[2:4]
            elif len(end) <= 2:
                # 只給了日，月份沿用區間開頭的月份
                mm = start[:2] if len(start) >= 3 else start[:1]
                dd = end
            else:
                mm, dd = end[:1], end[1:]
            return (int(mm), int(dd))
    except (ValueError, IndexError):
        return (0, 0)


def find_latest_115_file(folder_id):
    """列出資料夾內所有「歸仁日報表115」的檔案，用檔名日期(而非Drive建立時間)挑出真正最新的一份。
    這是必要的，因為回溯匯入歷史檔案時，Drive標記的建立時間是匯入當下的順序，
    不是檔名本身的日期，直接照createdTime排序會抓到錯的（例如抓到1月的舊檔）。"""
    from drive_reader import get_drive_service
    service = get_drive_service()
    query = (
        f"'{folder_id}' in parents and "
        f"name contains '歸仁日報表115' and trashed = false"
    )
    resp = service.files().list(
        q=query, pageSize=200, fields="files(id, name, modifiedTime)"
    ).execute()
    files = resp.get("files", [])
    if not files:
        return None
    files_with_date = [(f, _parse_filename_date(f["name"])) for f in files]
    files_with_date.sort(key=lambda x: x[1], reverse=True)
    return files_with_date[0][0]


def find_closest_year_file(folder_id, year_code, target_month, target_day):
    """在資料夾裡找「歸仁日報表{year_code}」的檔案，挑出檔名日期離 target_month/target_day
    最接近的一份（去年沒有正好同一天的檔案時，用最近的一天代替）。"""
    from drive_reader import get_drive_service
    service = get_drive_service()
    query = (
        f"'{folder_id}' in parents and "
        f"name contains '歸仁日報表{year_code}' and trashed = false"
    )
    resp = service.files().list(
        q=query, pageSize=300, fields="files(id, name)"
    ).execute()
    files = resp.get("files", [])
    if not files:
        return None

    def day_of_year(mm, dd):
        # 用月份*31+日 當簡化的年內序數，足夠拿來比大小找最近
        return mm * 31 + dd

    target_ord = day_of_year(target_month, target_day)
    best = None
    best_diff = None
    for f in files:
        mm, dd = _parse_filename_date(f["name"])
        if mm == 0:
            continue
        diff = abs(day_of_year(mm, dd) - target_ord)
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best = f
    return best


def parse_year_summary(wb, sheet_name):
    """解析「XXX年度」總表，回傳 {person: 年度累計領牌}（col27）。
    這個表所有課別的人都混在同一張表裡，直接掃描全部列即可，不用管課別區塊。"""
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    result = {}
    for r in range(1, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not is_person_name(name):
            continue
        name = str(name).strip()
        val = ws.cell(row=r, column=27).value
        if isinstance(val, (int, float)):
            result[name] = val
    return result


HISTORY_FILE = "dashboard_order_history.json"
SEED_HISTORY_FILE = "order_history_seed.json"


def load_order_history():
    history = {}
    if os.path.exists(SEED_HISTORY_FILE):
        with open(SEED_HISTORY_FILE, "r", encoding="utf-8") as f:
            history.update(json.load(f))
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            history.update(json.load(f))  # 累積的正式紀錄覆蓋種子資料的同日期項目
    return history


def save_order_history(history):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=1)


def compute_last_order_tracking(history, today_str):
    """掃描累積的每日訂單快照，算出每人每車型「最後一筆訂單」的日期與距今天數。
    注意：月累數字每個月初會重新歸零，所以「跨月」比較不能直接看數字增減
    （否則6/30某車型月累=4、7/4新月累=1會被誤判成「減少=沒有新訂單」，
    漏掉7月初其實已經有新訂單這件事）。做法：先把日期依「年-月」分組，
    只在同一個月份內部比較是否增加；每個月只要一開始(該月最早一筆快照)就有值，
    也視為一筆事件發生在那個日期（因為代表月初到那天之間已經有新訂單）。"""
    dates_sorted = sorted(history.keys())
    people = set()
    models = set()
    for snap in history.values():
        for p, md in snap.items():
            people.add(p)
            models.update(md.keys())

    # 依年月分組
    from collections import OrderedDict
    by_month = OrderedDict()
    for d in dates_sorted:
        ym = d[:7]  # "YYYY-MM"
        by_month.setdefault(ym, []).append(d)

    today = date.fromisoformat(today_str)
    result = {}
    for p in people:
        result[p] = {}
        for model in models:
            last_date = None
            for ym, ds in by_month.items():
                prev_val = None
                for d in ds:
                    cur_val = history[d].get(p, {}).get(model, 0)
                    if prev_val is None:
                        if cur_val > 0:
                            last_date = d  # 這個月一開始就有值，視為這個月初已經有新訂單
                    else:
                        if cur_val > prev_val:
                            last_date = d
                    prev_val = cur_val
            if last_date is None:
                continue
            days_since = (today - date.fromisoformat(last_date)).days
            result[p][model] = {"last_order_date": last_date, "days_since": days_since}
    return result
MONTHS = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']


def is_person_name(name):
    BLACKLIST_SUBSTR = ['合計', '月累', '課', '歸一', '歸二', '歸三', '歸仁', '公司']
    if not name or not str(name).strip():
        return False
    name = str(name).strip()
    if any(s in name for s in BLACKLIST_SUBSTR):
        return False
    if len(name) < 2 or len(name) > 5:
        return False
    return True


def normalize_model(name):
    if name is None:
        return None
    n = str(name).strip().upper().replace(' ', '').replace('\n', '')
    mapping = {
        'ZRV': 'ZRV', 'CR-V': 'CR-V', 'CRV': 'CR-V', 'HRV': 'HRV', 'HR-V': 'HRV',
        'FIT': 'FIT', 'CIVIC': 'CIVIC', 'ODYSSEY': 'ODYSSEY', 'PRELUDE': 'PRELUDE',
        'ACCORD': 'ACCORD', 'INSIGHT': 'INSIGHT', 'CR-Z': 'CR-Z', 'CRZ': 'CR-Z',
    }
    return mapping.get(n, n)


def find_groups(ws, header_row, sub_row, start_col, end_col_exclusive):
    """回傳每個車型群組 (model_name, start_col, cumcol)。cumcol 取群組內第一個「月累」欄，
    避免抓到後面殘留、無標籤的孤立欄位。"""
    starts = []
    for c in range(start_col, end_col_exclusive):
        v = ws.cell(row=header_row, column=c).value
        if v is not None and str(v).strip():
            starts.append((c, str(v).strip()))
    groups = []
    for i, (c, name) in enumerate(starts):
        next_c = starts[i + 1][0] if i + 1 < len(starts) else end_col_exclusive
        cumcol = None
        for cc in range(c, next_c):
            sv = ws.cell(row=sub_row, column=cc).value
            if sv is not None and '月累' in str(sv):
                cumcol = cc
                break
        if cumcol is not None:
            groups.append((normalize_model(name), c, cumcol))
    return groups


def parse_month_sheet(wb, sheet_name):
    """解析單一月份工作表，回傳 {person: {model: {'領牌':x, '訂單':y}}}"""
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    reg_groups = find_groups(ws, 4, 5, 2, 36)
    ord_groups = find_groups(ws, 4, 5, 49, 76)

    result = {}
    for r in range(6, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not is_person_name(name):
            continue
        name = str(name).strip()
        result.setdefault(name, {})
        for model, gc, cumcol in reg_groups:
            v = ws.cell(row=r, column=cumcol).value
            v = v if isinstance(v, (int, float)) else 0
            result[name].setdefault(model, {}).setdefault('領牌', 0)
            result[name][model]['領牌'] += v
        for model, gc, cumcol in ord_groups:
            v = ws.cell(row=r, column=cumcol).value
            v = v if isinstance(v, (int, float)) else 0
            result[name].setdefault(model, {}).setdefault('訂單', 0)
            result[name][model]['訂單'] += v
    return result


def build_dashboard_data():
    file_info = find_latest_115_file(DAILY_REPORT_FOLDER_ID)
    if not file_info:
        raise RuntimeError("Drive 資料夾裡找不到115年歸仁日報表檔案")

    content = download_file(file_info["id"])
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)

    today = date.today()
    current_month_key = f"{today.month}月"

    ytd = defaultdict(lambda: defaultdict(lambda: {'領牌': 0, '訂單': 0}))
    months_to_sum = [m for m in MONTHS if m in wb.sheetnames]
    for m in months_to_sum:
        r = parse_month_sheet(wb, m)
        for person, models in r.items():
            for model, vals in models.items():
                ytd[person][model]['領牌'] += vals.get('領牌', 0)
                ytd[person][model]['訂單'] += vals.get('訂單', 0)

    item1 = {p: sum(v['領牌'] for v in models.values()) for p, models in ytd.items()}
    item4 = {p: {mo: v['領牌'] for mo, v in models.items() if v['領牌'] > 0}
             for p, models in ytd.items()}

    month_progress = {}
    if current_month_key in wb.sheetnames:
        cur = parse_month_sheet(wb, current_month_key)
        for p, models in cur.items():
            month_progress[p] = {
                '訂單': sum(v.get('訂單', 0) for v in models.values()),
                '領牌': sum(v.get('領牌', 0) for v in models.values()),
            }

    # ---- 項目3：最後訂單追蹤（累積每日快照後自動計算）----
    today_str = today.isoformat()
    history = load_order_history()
    if current_month_key in wb.sheetnames:
        cur_orders = parse_month_sheet(wb, current_month_key)
        snapshot = {p: {m: v.get('訂單', 0) for m, v in models.items() if '訂單' in v}
                    for p, models in cur_orders.items()}
        history[today_str] = snapshot
        save_order_history(history)
    last_order_tracking = compute_last_order_tracking(history, today_str)

    # ---- 項目2：跟去年同期比較（找去年同一天，沒有就找最接近的一天）----
    yoy_comparison = None
    try:
        last_year_file = find_closest_year_file(DAILY_REPORT_FOLDER_ID, "114", today.month, today.day)
        if last_year_file:
            ly_content = download_file(last_year_file["id"])
            ly_wb = openpyxl.load_workbook(io.BytesIO(ly_content), data_only=True)
            ly_item1 = parse_year_summary(ly_wb, "114年度")
            ly_month, ly_day = _parse_filename_date(last_year_file["name"])
            yoy_comparison = {
                "last_year_file": last_year_file["name"],
                "last_year_date": f"2025-{ly_month:02d}-{ly_day:02d}" if ly_month else None,
                "last_year_ytd": ly_item1,
            }
    except Exception as e:
        yoy_comparison = {"error": str(e)}

    data = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "source_file": file_info["name"],
        "item1_ytd_registration": item1,
        "item4_ytd_by_model": item4,
        "month_progress": month_progress,
        "last_order_tracking": last_order_tracking,
        "yoy_comparison": yoy_comparison,
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return data


if __name__ == "__main__":
    result = build_dashboard_data()
    print(json.dumps(result, ensure_ascii=False, indent=2))
