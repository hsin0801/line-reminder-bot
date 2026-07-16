"""
歸仁儀表板 - 資料解析模組（v3）
沿用現有 drive_reader.py 的 Drive 連線邏輯，不需要新的環境變數。

重要效能修正：
openpyxl 的 read_only=True 模式是為了「循序」讀取設計的（一行一行從頭讀到尾）。
之前版本用 ws.cell(row=r, column=c) 這種「跳著讀某一格」的隨機存取方式，
在 read_only 模式下，每讀一格都會重新掃描一次底層XML，導致比一般模式更慢、
更耗記憶體，最後被 Render 判定 OOM (Out of Memory) 強制關閉。

修正做法：每張工作表只用 ws.iter_rows(values_only=True) 循序讀「一次」，
轉成一個小小的 list-of-tuples（grid），之後所有查詢都對這個 grid 做操作，
不再直接碰 worksheet 物件。grid 本身很小（頂多百來行、兩百欄的數字/文字），
記憶體佔用可忽略不計。
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
MONTHS = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']


# ============ 檔名日期解析 ============

def _parse_filename_date(filename):
    """從檔名解析出這份報表對應的最後日期，回傳 (month, day) 供排序比較。
    年度前綴用 \\d{3} 泛用比對（113/114/115皆可通用）。
    解析失敗回傳 (0, 0)（排最後面，不會被誤判為最新）。
    """
    m = re.search(r'\d{3}\s*([0-9\-\s]+)', filename)
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
            if len(digits) == 3:
                mm, dd = digits[0], digits[1:]
            else:
                mm, dd = digits[:2], digits[2:4]
            return (int(mm), int(dd))
        else:
            end = parts[-1]
            start = parts[0]
            if len(end) >= 4:
                mm, dd = end[:2], end[2:4]
            elif len(end) <= 2:
                mm = start[:2] if len(start) >= 3 else start[:1]
                dd = end
            else:
                mm, dd = end[:1], end[1:]
            return (int(mm), int(dd))
    except (ValueError, IndexError):
        return (0, 0)


def find_latest_115_file(folder_id):
    """列出資料夾內所有「歸仁日報表115」的檔案，用檔名日期(而非Drive建立時間)挑出真正最新的一份。"""
    from drive_reader import get_drive_service
    service = get_drive_service()
    query = (
        f"'{folder_id}' in parents and "
        f"name contains '歸仁日報表115' and trashed = false"
    )
    resp = service.files().list(
        q=query, pageSize=200, fields="files(id, name)"
    ).execute()
    files = resp.get("files", [])
    if not files:
        return None
    files_with_date = [(f, _parse_filename_date(f["name"])) for f in files]
    files_with_date.sort(key=lambda x: x[1], reverse=True)
    return files_with_date[0][0]


def find_closest_114_file(folder_id, target_month, target_day):
    """在資料夾裡找「歸仁日報表114」的檔案，挑出檔名日期離 target_month/target_day 最接近的一份。"""
    from drive_reader import get_drive_service
    service = get_drive_service()
    query = (
        f"'{folder_id}' in parents and "
        f"name contains '歸仁日報表114' and trashed = false"
    )
    resp = service.files().list(
        q=query, pageSize=300, fields="files(id, name)"
    ).execute()
    files = resp.get("files", [])
    if not files:
        return None, "Drive查詢114年檔案回傳空清單"

    def day_ord(mm, dd):
        return mm * 31 + dd

    target_ord = day_ord(target_month, target_day)
    files_with_date = [(f, _parse_filename_date(f["name"])) for f in files]
    files_with_date = [(f, md) for f, md in files_with_date if md[0] != 0]
    if not files_with_date:
        return None, "114年檔案都解析不出日期: " + ",".join(f["name"] for f in files[:5])

    files_with_date.sort(key=lambda x: abs(day_ord(*x[1]) - target_ord))
    best_file, (mm, dd) = files_with_date[0]
    return (best_file, mm, dd), None


# ============ Grid：把工作表一次循序讀成小陣列 ============

def sheet_to_grid(wb, sheet_name, max_cols=210):
    """把整張工作表用循序方式(iter_rows)讀一次，轉成 list of tuples。
    之後所有查詢都對這個 grid 操作，不再用 ws.cell() 隨機存取。"""
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    grid = []
    for row in ws.iter_rows(min_col=1, max_col=max_cols, values_only=True):
        grid.append(row)
    return grid


def grid_get(grid, row, col):
    """grid裡取值，row/col都是1-indexed（配合Excel習慣），超出範圍回傳None。"""
    r_idx = row - 1
    c_idx = col - 1
    if r_idx < 0 or r_idx >= len(grid):
        return None
    row_tuple = grid[r_idx]
    if c_idx < 0 or c_idx >= len(row_tuple):
        return None
    return row_tuple[c_idx]


# ============ 共用小工具 ============

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


def find_groups(grid, header_row, sub_row, start_col, end_col_exclusive):
    """回傳每個車型群組 (model_name, start_col, cumcol)。cumcol 取群組內第一個「月累」欄。"""
    starts = []
    for c in range(start_col, end_col_exclusive):
        v = grid_get(grid, header_row, c)
        if v is not None and str(v).strip():
            starts.append((c, str(v).strip()))
    groups = []
    for i, (c, name) in enumerate(starts):
        next_c = starts[i + 1][0] if i + 1 < len(starts) else end_col_exclusive
        cumcol = None
        for cc in range(c, next_c):
            sv = grid_get(grid, sub_row, cc)
            if sv is not None and '月累' in str(sv):
                cumcol = cc
                break
        if cumcol is not None:
            groups.append((normalize_model(name), c, cumcol))
    return groups


# ============ 月份工作表解析（訂單/領牌，車型層級）============

def parse_month_sheet(wb, sheet_name):
    """解析單一月份工作表，回傳 {person: {model: {'領牌':x, '訂單':y}}}"""
    grid = sheet_to_grid(wb, sheet_name, max_cols=80)
    if not grid:
        return {}
    reg_groups = find_groups(grid, 4, 5, 2, 36)
    ord_groups = find_groups(grid, 4, 5, 49, 76)

    result = {}
    for r in range(6, len(grid) + 1):
        name = grid_get(grid, r, 1)
        if not is_person_name(name):
            continue
        name = str(name).strip()
        result.setdefault(name, {})
        for model, gc, cumcol in reg_groups:
            v = grid_get(grid, r, cumcol)
            v = v if isinstance(v, (int, float)) else 0
            result[name].setdefault(model, {}).setdefault('領牌', 0)
            result[name][model]['領牌'] += v
        for model, gc, cumcol in ord_groups:
            v = grid_get(grid, r, cumcol)
            v = v if isinstance(v, (int, float)) else 0
            result[name].setdefault(model, {}).setdefault('訂單', 0)
            result[name][model]['訂單'] += v
    return result


# ============ 年度總表解析（只要col27年度累計領牌）============

def parse_year_summary(wb, sheet_name):
    """解析「XXX年度」總表，回傳 {person: 年度累計領牌}（col27）。"""
    grid = sheet_to_grid(wb, sheet_name, max_cols=30)
    if not grid:
        return {}
    result = {}
    for r in range(1, len(grid) + 1):
        name = grid_get(grid, r, 1)
        if not is_person_name(name):
            continue
        name = str(name).strip()
        val = grid_get(grid, r, 27)
        if isinstance(val, (int, float)):
            result[name] = val
    return result


# ============ 項目3：每日訂單快照歷史 ============

HISTORY_FILE = "dashboard_order_history.json"
SEED_HISTORY_FILE = "order_history_seed.json"


def load_order_history():
    history = {}
    if os.path.exists(SEED_HISTORY_FILE):
        with open(SEED_HISTORY_FILE, "r", encoding="utf-8") as f:
            history.update(json.load(f))
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            history.update(json.load(f))
    return history


def save_order_history(history):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=1)


def compute_last_order_tracking(history, today_str):
    """掃描累積的每日訂單快照，算出每人每車型「最後一筆訂單」的日期與距今天數。"""
    dates_sorted = sorted(history.keys())
    people = set()
    models = set()
    for snap in history.values():
        for p, md in snap.items():
            people.add(p)
            models.update(md.keys())

    from collections import OrderedDict
    by_month = OrderedDict()
    for d in dates_sorted:
        ym = d[:7]
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
                            last_date = d
                    else:
                        if cur_val > prev_val:
                            last_date = d
                    prev_val = cur_val
            if last_date is None:
                continue
            days_since = (today - date.fromisoformat(last_date)).days
            result[p][model] = {"last_order_date": last_date, "days_since": days_since}
    return result


# ============ 主流程 ============

def build_dashboard_data():
    file_info = find_latest_115_file(DAILY_REPORT_FOLDER_ID)
    if not file_info:
        raise RuntimeError("Drive 資料夾裡找不到115年歸仁日報表檔案")

    content = download_file(file_info["id"])
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)

    today = date.today()
    current_month_key = f"{today.month}月"

    ytd = defaultdict(lambda: defaultdict(lambda: {'領牌': 0, '訂單': 0}))
    months_to_sum = [m for m in MONTHS if m in wb.sheetnames]
    month_sheets_cache = {}
    for m in months_to_sum:
        r = parse_month_sheet(wb, m)
        month_sheets_cache[m] = r
        for person, models in r.items():
            for model, vals in models.items():
                ytd[person][model]['領牌'] += vals.get('領牌', 0)
                ytd[person][model]['訂單'] += vals.get('訂單', 0)

    item1 = {p: sum(v['領牌'] for v in models.values()) for p, models in ytd.items()}
    item4 = {p: {mo: v['領牌'] for mo, v in models.items() if v['領牌'] > 0}
             for p, models in ytd.items()}

    cur = month_sheets_cache.get(current_month_key, {})
    month_progress = {}
    for p, models in cur.items():
        month_progress[p] = {
            '訂單': sum(v.get('訂單', 0) for v in models.values()),
            '領牌': sum(v.get('領牌', 0) for v in models.values()),
        }

    # ---- 項目3：最後訂單追蹤（累積每日快照後自動計算）----
    today_str = today.isoformat()
    history = load_order_history()
    if cur:
        snapshot = {p: {m: v.get('訂單', 0) for m, v in models.items() if '訂單' in v}
                    for p, models in cur.items()}
        history[today_str] = snapshot
        save_order_history(history)
    last_order_tracking = compute_last_order_tracking(history, today_str)

    wb.close()
    del content

    # ---- 項目2：跟去年同期比較（找去年同一天，沒有就找最接近的一天）----
    yoy_comparison = None
    try:
        found, err = find_closest_114_file(DAILY_REPORT_FOLDER_ID, today.month, today.day)
        if found is None:
            yoy_comparison = {"error": err}
        else:
            last_year_file, ly_month, ly_day = found
            ly_content = download_file(last_year_file["id"])
            ly_wb = openpyxl.load_workbook(io.BytesIO(ly_content), data_only=True, read_only=True)
            ly_item1 = parse_year_summary(ly_wb, "114年度")
            ly_wb.close()
            del ly_content
            yoy_comparison = {
                "last_year_file": last_year_file["name"],
                "last_year_date": f"2025-{ly_month:02d}-{ly_day:02d}",
                "last_year_ytd": ly_item1,
            }
    except Exception as e:
        import traceback
        yoy_comparison = {"error": str(e), "trace": traceback.format_exc()[-500:]}

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
