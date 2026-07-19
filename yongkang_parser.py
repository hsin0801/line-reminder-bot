"""
永康儀表板 - 資料解析模組
永康日報表是舊版 .xls 格式，用 xlrd 讀取（不是 openpyxl）。
欄位結構、課別編制都跟歸仁不同，所以是獨立的一份解析程式，
但整體「找最新檔案 → 解析月份工作表 → 累加年度 → 課別小計 → 歷史快照」的
架構跟 dashboard_parser.py(歸仁) 是一樣的邏輯，方便之後維護對照。
"""

import io
import os
import re
import json
from datetime import date, datetime
from collections import defaultdict, OrderedDict

import xlrd

from drive_reader import download_file, DAILY_REPORT_FOLDER_ID

DATA_FILE = "yongkang_dashboard_data.json"
MONTHS = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']

# 只追蹤這幾個車型，其餘(CITY/ODYSSEY等)忽略
TRACKED_MODELS = {'CR-V', 'HRV', 'FIT', 'CIVIC', 'PRELUDE'}

TEAM_STRUCTURE = {
    '永康一課': ['謝岱宏', '林昭吾', '林明昌', '陳靖玟', '林駿軒'],
    '永康二課': ['洪嘉綺', '林祐諄', '王少佟', '李霽杰', '陳冠甫'],
    '永康三課': ['蘇嵩閔', '陳建志', '黃湘婷', '江念澤', '王裕維', '張詠竣'],
}
PERSON_TO_DEPT = {p: dept for dept, members in TEAM_STRUCTURE.items() for p in members}
TEAM_ORDER = [p for members in TEAM_STRUCTURE.values() for p in members]

BLACKLIST_SUBSTR = ['合計', '月累', '課', '營業', '永康', '公司', '領牌', '訂單']


def is_person_name(name):
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
        'CR-V': 'CR-V', 'CRV': 'CR-V', 'HR-V': 'HRV', 'HRV': 'HRV',
        'FIT': 'FIT', 'CITY': 'CITY', 'CIVIC': 'CIVIC',
        'PRELUDE': 'PRELUDE', 'ODYSSEY': 'ODYSSEY', 'ACCORD': 'ACCORD',
    }
    return mapping.get(n, n)


# ============ 檔名日期解析（沿用跟歸仁一樣的邏輯）============

def _parse_filename_date(filename):
    m = re.search(r'\d{3}\s*([0-9\-\s]+)', filename)
    if not m:
        return (0, 0)
    raw = m.group(1).replace(' ', '').strip('-')
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
    from drive_reader import get_drive_service
    service = get_drive_service()
    query = f"'{folder_id}' in parents and name contains '永康日報表115' and trashed = false"
    resp = service.files().list(q=query, pageSize=300, fields="files(id, name)").execute()
    files = resp.get("files", [])
    if not files:
        return None
    files_with_date = [(f, _parse_filename_date(f["name"])) for f in files]
    files_with_date.sort(key=lambda x: x[1], reverse=True)
    return files_with_date[0][0]


def find_closest_114_file(folder_id, target_month, target_day):
    from drive_reader import get_drive_service
    service = get_drive_service()
    query = f"'{folder_id}' in parents and name contains '永康日報表114' and trashed = false"
    resp = service.files().list(q=query, pageSize=300, fields="files(id, name)").execute()
    files = resp.get("files", [])
    if not files:
        return None, "Drive查詢114年永康檔案回傳空清單"

    def day_ord(mm, dd):
        return mm * 31 + dd

    target_ord = day_ord(target_month, target_day)
    files_with_date = [(f, _parse_filename_date(f["name"])) for f in files]
    files_with_date = [(f, md) for f, md in files_with_date if md[0] != 0]
    if not files_with_date:
        return None, "114年永康檔案都解析不出日期"
    files_with_date.sort(key=lambda x: abs(day_ord(*x[1]) - target_ord))
    best_file, (mm, dd) = files_with_date[0]
    return (best_file, mm, dd), None


# ============ xlrd Grid 工具 ============

def sheet_to_grid(wb, sheet_name):
    if sheet_name not in wb.sheet_names():
        return []
    sh = wb.sheet_by_name(sheet_name)
    grid = []
    for r in range(sh.nrows):
        grid.append(tuple(sh.cell_value(r, c) for c in range(sh.ncols)))
    return grid


def grid_get(grid, row, col):
    r_idx, c_idx = row - 1, col - 1
    if r_idx < 0 or r_idx >= len(grid):
        return None
    row_tuple = grid[r_idx]
    if c_idx < 0 or c_idx >= len(row_tuple):
        return None
    v = row_tuple[c_idx]
    return v if v != '' else None


def find_groups(grid, header_row, sub_row, start_col, end_col_exclusive):
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
            model = normalize_model(name)
            if model in TRACKED_MODELS:
                groups.append((model, c, cumcol))
    return groups


# ============ 月份工作表解析 ============
# 領牌區塊: 1-idx col2~36 (對應xlrd 0-idx col1~35)；訂單區塊: 1-idx col54~87 (對應xlrd 0-idx col53~86)
# 表頭列: header_row=3 (xlrd 0-idx row2)，trim列: sub_row=4 (xlrd 0-idx row3)

def parse_month_sheet(wb, sheet_name):
    grid = sheet_to_grid(wb, sheet_name)
    if not grid:
        return {}
    reg_groups = find_groups(grid, 3, 4, 2, 37)
    ord_groups = find_groups(grid, 3, 4, 54, 87)

    result = {}
    for r in range(1, len(grid) + 1):
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


# CR-V配備等級細分（只用於項目3「最後一筆訂單追蹤」，不影響項目1/4的車型層級統計）
# 使用者已確認的固定欄位定義(1-indexed，對應xlrd 0-indexed 58/59/60/61)：
#   永康訂單區塊 CR-V：col59-60=PET(VTI-S/S)，col61-62=e:HEV(ES/EP)
# 注意：配備等級細分欄位常常沒填(只填車型層級總計)，這裡只算「有填」的部分，
# 沒填的視為沒有該等級訂單(使用者已確認接受這個取捨)。
CRV_TRIM_ORDER_COLS = {'CR-V(PET)': [59, 60], 'CR-V(e:HEV)': [61, 62]}


def parse_crv_trim_split(wb, sheet_name):
    """解析CR-V配備等級細分的訂單數，回傳 {person: {'CR-V(PET)':x, 'CR-V(e:HEV)':y}}"""
    grid = sheet_to_grid(wb, sheet_name)
    if not grid:
        return {}
    result = {}
    for r in range(1, len(grid) + 1):
        name = grid_get(grid, r, 1)
        if not is_person_name(name):
            continue
        name = str(name).strip()
        result.setdefault(name, {})
        for label, cols in CRV_TRIM_ORDER_COLS.items():
            total = 0
            for c in cols:
                v = grid_get(grid, r, c)
                if isinstance(v, (int, float)):
                    total += v
            result[name][label] = total
    return result


def parse_year_summary(wb, sheet_name):
    """解析「XXX年度累計報表」總表，回傳{person: 年度累計領牌}。
    已核對實際檔案：跟歸仁一樣，年度累計領牌固定在col27(1-indexed)。"""
    grid = sheet_to_grid(wb, sheet_name)
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
            result[name] = result.get(name, 0) + val
    return result


# ============ 課別小計 ============

def compute_dept_totals_scalar(person_value_dict):
    totals = {dept: 0 for dept in TEAM_STRUCTURE}
    for p, v in person_value_dict.items():
        dept = PERSON_TO_DEPT.get(p)
        if dept:
            totals[dept] += v
    return totals


def compute_dept_totals_by_model(person_model_dict):
    totals = {dept: defaultdict(int) for dept in TEAM_STRUCTURE}
    for p, models in person_model_dict.items():
        dept = PERSON_TO_DEPT.get(p)
        if dept:
            for model, v in models.items():
                totals[dept][model] += v
    return {dept: dict(models) for dept, models in totals.items()}


def sort_by_total_desc(d):
    return dict(sorted(d.items(), key=lambda item: sum(item[1].values()) if isinstance(item[1], dict) else item[1], reverse=True))


# ============ 項目3：每日訂單快照歷史 ============

# ============ 項目3：每日訂單快照歷史 ============
# 改成存在 Google Drive（不是本機硬碟），理由同歸仁那份：Render免費方案
# 本機硬碟會在服務休眠/重啟時被清空，存本機的話累積的歷史會一直不見。

HISTORY_DRIVE_FILENAME = "yongkang_order_history.json"
SEED_HISTORY_FILE = "yongkang_order_history_seed.json"  # 隨程式碼一起部署的靜態種子檔


def load_order_history():
    from drive_json_store import load_json_from_drive

    history = {}
    if os.path.exists(SEED_HISTORY_FILE):
        with open(SEED_HISTORY_FILE, "r", encoding="utf-8") as f:
            history.update(json.load(f))
    history.update(load_json_from_drive(DAILY_REPORT_FOLDER_ID, HISTORY_DRIVE_FILENAME))

    # 遷移清理：改成CR-V拆分成PET/e:HEV之前，Drive上可能已經存了舊格式的
    # 合併「CR-V」欄位，留著會跟新的拆分欄位一起出現變成三欄，這裡讀取時
    # 順手清掉，讓資料統一都是拆分後的格式。
    for snap in history.values():
        for models in snap.values():
            models.pop('CR-V', None)

    return history


def save_order_history(history):
    from drive_json_store import save_json_to_drive

    save_json_to_drive(DAILY_REPORT_FOLDER_ID, HISTORY_DRIVE_FILENAME, history)


MONTH_LAST_DAY_2026 = {
    '1月': (1, 31), '2月': (2, 28), '3月': (3, 31), '4月': (4, 30),
    '5月': (5, 31), '6月': (6, 30), '7月': (7, 31), '8月': (8, 31),
    '9月': (9, 30), '10月': (10, 31), '11月': (11, 30), '12月': (12, 31),
}


DAILY_VALUE_MODELS = {'CR-V(PET)', 'CR-V(e:HEV)'}  # 這幾個欄位是「當天有填=當天有訂單」，不是月累計


def compute_last_order_tracking(history, today_str):
    dates_sorted = sorted(history.keys())
    people, models = set(), set()
    for snap in history.values():
        for p, md in snap.items():
            people.add(p)
            models.update(md.keys())

    by_month = OrderedDict()
    for d in dates_sorted:
        by_month.setdefault(d[:7], []).append(d)

    # 歷史紀錄只有一筆快照時（例如剛開始追蹤的第一天），「這個月第一筆就有數字」
    # 這個判斷完全沒有意義——沒有更早的資料可以比較，沒辦法知道訂單確切是哪一天下的，
    # 只知道「月初到現在之間某天有」。這種情況全部交給 fill_fallback_from_monthly
    # 用月度封存退回近似值處理，這裡不產生任何「精確日期」的結果，避免誤判成「0天」。
    # (這個限制只適用於「月累計」欄位；CR-V(PET)/CR-V(e:HEV)是當日值欄位，不受此限制。)
    global_first_date = dates_sorted[0] if dates_sorted else None
    only_one_snapshot_ever = len(dates_sorted) <= 1

    today = date.fromisoformat(today_str)
    result = {}
    for p in people:
        result[p] = {}
        for model in models:
            last_date = None
            if model in DAILY_VALUE_MODELS:
                # 當日值欄位：直接找最近一次「當天有填數字」的日期，不受單一快照限制
                for d in dates_sorted:
                    cur_val = history[d].get(p, {}).get(model, 0)
                    if cur_val > 0:
                        last_date = d
            else:
                if only_one_snapshot_ever:
                    continue
                for ym, ds in by_month.items():
                    prev_val = None
                    for d in ds:
                        cur_val = history[d].get(p, {}).get(model, 0)
                        if prev_val is None:
                            if cur_val > 0 and d != global_first_date:
                                last_date = d
                        else:
                            if cur_val > prev_val:
                                last_date = d
                        prev_val = cur_val
            if last_date is None:
                continue
            days_since = (today - date.fromisoformat(last_date)).days
            result[p][model] = {"last_order_date": last_date, "days_since": days_since, "approx": False}
    return result


def fill_fallback_from_monthly(last_order_tracking, month_sheets_cache, months_to_sum, today_str):
    """歷史快照剛開始累積，早期資料少，用月度封存表退回月底當近似值。"""
    today = date.fromisoformat(today_str)
    people_models_with_data = set()
    for m in months_to_sum:
        for p, models in month_sheets_cache.get(m, {}).items():
            for model, vals in models.items():
                if vals.get('訂單', 0) > 0:
                    people_models_with_data.add((p, model))

    for p, model in people_models_with_data:
        if model in DAILY_VALUE_MODELS:
            continue  # 當日值欄位，月度封存退回機制對它沒有意義
        if last_order_tracking.get(p, {}).get(model) is not None:
            continue
        for m in reversed(months_to_sum):
            v = month_sheets_cache.get(m, {}).get(p, {}).get(model, {}).get('訂單', 0)
            if v > 0:
                mm, dd = MONTH_LAST_DAY_2026[m]
                candidate = date(today.year, mm, dd)
                if candidate > today:
                    candidate = today
                last_order_tracking.setdefault(p, {})
                last_order_tracking[p][model] = {
                    "last_order_date": candidate.isoformat(),
                    "days_since": (today - candidate).days,
                    "approx": True,
                }
                break
    return last_order_tracking


# ============ 主流程 ============

def build_daily_snapshot(wb, month_key):
    """組出跟每日正常流程一樣格式的快照：非CR-V車型用月累計(群組層級)，
    CR-V拆成PET/e:HEV用當日值。"""
    r = parse_month_sheet(wb, month_key)
    crv_split = parse_crv_trim_split(wb, month_key)
    snapshot = {}
    for person, models in r.items():
        snapshot[person] = {k: v.get('訂單', 0) for k, v in models.items()
                             if '訂單' in v and k != 'CR-V'}
    for person, split_vals in crv_split.items():
        snapshot.setdefault(person, {})
        for label, val in split_vals.items():
            snapshot[person][label] = val
    return snapshot


def backfill_full_history(max_seconds=240):
    """完整回溯歷史：把Drive資料夾裡所有115年永康日報表逐一解析，
    重建出每一天的訂單快照(含CR-V PET/e:HEV當日值)，寫入Drive上的歷史檔案。
    可分批執行，已處理過的日期會跳過，重複呼叫可接續進度直到全部處理完。"""
    import time
    start_time = time.time()

    from drive_reader import get_drive_service, download_file
    service = get_drive_service()
    query = f"'{DAILY_REPORT_FOLDER_ID}' in parents and name contains '永康日報表115' and trashed = false"
    resp = service.files().list(q=query, pageSize=300, fields="files(id, name)").execute()
    files = resp.get("files", [])

    files_with_date = [(f, _parse_filename_date(f["name"])) for f in files]
    files_with_date = [(f, md) for f, md in files_with_date if md[0] != 0]
    files_with_date.sort(key=lambda x: x[1])

    history = load_order_history()
    processed = 0
    errors = []
    timed_out = False

    for f, (mm, dd) in files_with_date:
        date_str = f"2026-{mm:02d}-{dd:02d}"
        if date_str in history:
            continue
        if time.time() - start_time > max_seconds:
            timed_out = True
            break
        try:
            content = download_file(f["id"])
            wb = xlrd.open_workbook(file_contents=content)
            month_key = f"{mm}月"
            if month_key in wb.sheet_names():
                history[date_str] = build_daily_snapshot(wb, month_key)
                processed += 1
            del content
        except Exception as e:
            errors.append(f"{f['name']}: {str(e)[:100]}")

    save_order_history(history)

    return {
        "processed_this_run": processed,
        "total_files_found": len(files_with_date),
        "dates_now_in_history": len(history),
        "timed_out": timed_out,
        "done": not timed_out,
        "errors": errors[:10],
    }


def build_yongkang_data():
    file_info = find_latest_115_file(DAILY_REPORT_FOLDER_ID)
    if not file_info:
        raise RuntimeError("Drive 資料夾裡找不到115年永康日報表檔案")

    content = download_file(file_info["id"])
    wb = xlrd.open_workbook(file_contents=content)

    today = date.today()
    current_month_key = f"{today.month}月"

    ytd = defaultdict(lambda: defaultdict(lambda: {'領牌': 0, '訂單': 0}))
    months_to_sum = [m for m in MONTHS if m in wb.sheet_names()]
    month_sheets_cache = {}
    month_sheets_cache_for_tracking = {}  # CR-V拆成PET/e:HEV後的版本，只給項目3用
    for m in months_to_sum:
        r = parse_month_sheet(wb, m)
        month_sheets_cache[m] = r
        for person, models in r.items():
            for model, vals in models.items():
                ytd[person][model]['領牌'] += vals.get('領牌', 0)
                ytd[person][model]['訂單'] += vals.get('訂單', 0)

        crv_split = parse_crv_trim_split(wb, m)
        r_tracking = {}
        for person, models in r.items():
            r_tracking[person] = {k: v for k, v in models.items() if k != 'CR-V'}
        for person, split_vals in crv_split.items():
            r_tracking.setdefault(person, {})
            for label, val in split_vals.items():
                r_tracking[person][label] = {'訂單': val}
        month_sheets_cache_for_tracking[m] = r_tracking

    item1 = {p: sum(v['領牌'] for v in models.values()) for p, models in ytd.items()}
    team_total_ytd_registration = sum(item1.values())
    item1_dept_totals = compute_dept_totals_scalar(item1)

    item4 = {p: {mo: v['領牌'] for mo, v in models.items() if v['領牌'] > 0}
             for p, models in ytd.items()}
    item4 = sort_by_total_desc(item4)
    item4_dept_totals = compute_dept_totals_by_model(item4)

    cur = month_sheets_cache.get(current_month_key, {})
    cur_tracking = month_sheets_cache_for_tracking.get(current_month_key, {})
    month_progress = {}
    for p, models in cur.items():
        month_progress[p] = {
            '訂單': sum(v.get('訂單', 0) for v in models.values()),
            '領牌': sum(v.get('領牌', 0) for v in models.values()),
        }
    month_progress_dept_totals = compute_dept_totals_by_model(month_progress)

    today_str = today.isoformat()
    history = load_order_history()
    if cur_tracking:
        snapshot = {p: {m: v.get('訂單', 0) for m, v in models.items() if '訂單' in v}
                    for p, models in cur_tracking.items()}
        history[today_str] = snapshot
        save_order_history(history)
    last_order_tracking = compute_last_order_tracking(history, today_str)
    last_order_tracking = fill_fallback_from_monthly(
        last_order_tracking, month_sheets_cache_for_tracking, months_to_sum, today_str
    )

    # 不在目前三課名單裡的人（如已離職的華品如/洪立綱/林雯秀等），個人列表不顯示，
    # 但課別小計/團隊總量在前面已經算過了(item1_dept_totals, team_total_ytd_registration)，不受影響
    item1 = {p: v for p, v in item1.items() if p in TEAM_ORDER}
    item4 = {p: v for p, v in item4.items() if p in TEAM_ORDER}
    month_progress = {p: v for p, v in month_progress.items() if p in TEAM_ORDER}
    last_order_tracking = {p: v for p, v in last_order_tracking.items() if p in TEAM_ORDER}

    # ---- 去年同期比較（找去年同一天，沒有就找最接近的一天）----
    yoy_comparison = None
    try:
        found, err = find_closest_114_file(DAILY_REPORT_FOLDER_ID, today.month, today.day)
        if found is None:
            yoy_comparison = {"error": err}
        else:
            last_year_file, ly_month, ly_day = found
            ly_content = download_file(last_year_file["id"])
            ly_wb = xlrd.open_workbook(file_contents=ly_content)
            ly_sheet_name = next((s for s in ly_wb.sheet_names() if '年度累計報表' in s), None)
            if ly_sheet_name is None:
                yoy_comparison = {"error": "114年檔案裡找不到「年度累計報表」工作表",
                                   "sheet_names": ly_wb.sheet_names()}
            else:
                ly_item1 = parse_year_summary(ly_wb, ly_sheet_name)
                yoy_comparison = {
                    "last_year_file": last_year_file["name"],
                    "last_year_date": f"2025-{ly_month:02d}-{ly_day:02d}",
                    "last_year_ytd": {k: int(v) for k, v in ly_item1.items()},
                }
    except Exception as e:
        import traceback
        yoy_comparison = {"error": str(e), "trace": traceback.format_exc()[-500:]}

    yoy_last_year_dept_totals = None
    yoy_this_year_dept_totals = None
    if yoy_comparison and yoy_comparison.get("last_year_ytd") is not None:
        yoy_last_year_dept_totals = compute_dept_totals_scalar(yoy_comparison["last_year_ytd"])
        yoy_this_year_dept_totals = item1_dept_totals

    # ---- 全部轉成 int，台數不需要小數點（xlrd讀出來的數字預設是float）----
    def to_int(x):
        try:
            return int(round(x))
        except (TypeError, ValueError):
            return x

    def clean_scalar_dict(d):
        return {k: to_int(v) for k, v in d.items()}

    def clean_model_dict(d):
        return {k: {m: to_int(v) for m, v in models.items()} for k, models in d.items()}

    item1 = clean_scalar_dict(item1)
    item1_dept_totals = clean_scalar_dict(item1_dept_totals)
    team_total_ytd_registration = to_int(team_total_ytd_registration)
    item4 = clean_model_dict(item4)
    item4_dept_totals = clean_model_dict(item4_dept_totals)
    month_progress = clean_model_dict(month_progress)
    month_progress_dept_totals = clean_model_dict(month_progress_dept_totals)
    if yoy_comparison and yoy_comparison.get("last_year_ytd") is not None:
        yoy_last_year_dept_totals = clean_scalar_dict(yoy_last_year_dept_totals)
        yoy_this_year_dept_totals = item1_dept_totals

    data = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "source_file": file_info["name"],
        "team_structure": TEAM_STRUCTURE,
        "item1_ytd_registration": item1,
        "item1_dept_totals": item1_dept_totals,
        "team_total_ytd_registration": team_total_ytd_registration,
        "item4_ytd_by_model": item4,
        "item4_dept_totals": item4_dept_totals,
        "month_progress": month_progress,
        "month_progress_dept_totals": month_progress_dept_totals,
        "last_order_tracking": last_order_tracking,
        "yoy_comparison": yoy_comparison,
        "yoy_last_year_dept_totals": yoy_last_year_dept_totals,
        "yoy_this_year_dept_totals": yoy_this_year_dept_totals,
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return data


if __name__ == "__main__":
    result = build_yongkang_data()
    print(json.dumps(result, ensure_ascii=False, indent=2))
