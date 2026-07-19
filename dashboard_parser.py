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


# CR-V配備等級細分（只用於項目3「最後一筆訂單追蹤」，不影響項目1/4的車型層級統計）
# 使用者已確認的固定欄位定義：
#   歸仁訂單區塊 CR-V：col49-50=PET(1.5vti-s/S)，col51-52=e:HEV(e:HEVS/e:HEVP)
# 注意：配備等級細分欄位常常沒填(只填車型層級總計)，這裡只算「有填」的部分，
# 沒填的視為沒有該等級訂單(使用者已確認接受這個取捨)。
CRV_TRIM_ORDER_COLS = {'CR-V(PET)': [49, 50], 'CR-V(e:HEV)': [51, 52]}


def parse_crv_trim_split(wb, sheet_name):
    """解析CR-V配備等級細分的訂單數，回傳 {person: {'CR-V(PET)':x, 'CR-V(e:HEV)':y}}

    重要：CR-V的欄位命名今年年中換過。1~4月訂單區塊col49-52是「2.0vti/1.5vti-s/1.5S/1.5P」
    這種舊款命名(完全不是PET/e:HEV)，5月起才變成「1.5vti-s/S/e:HEVS/e:HEVP」(真正的PET/e:HEV)。
    如果不驗證直接套用固定欄位，1~4月的舊款配備會被誤判成PET/e:HEV，
    產生錯誤的歷史訂單紀錄，這裡先確認欄位標籤符合預期才抓取，不符合就跳過該月，
    避免不同月份的欄位定義互相污染。"""
    grid = sheet_to_grid(wb, sheet_name, max_cols=80)
    if not grid:
        return {}

    # 驗證col51/52(e:HEV欄位位置)的標籤是不是真的含有'HEV'字樣，
    # 不是就代表這個月是舊款欄位定義，整個跳過不抓取(回傳空dict)
    sub_row = 5
    label_51 = grid_get(grid, sub_row, 51)
    label_52 = grid_get(grid, sub_row, 52)
    is_new_layout = (
        label_51 and 'HEV' in str(label_51).upper()
        and label_52 and 'HEV' in str(label_52).upper()
    )
    if not is_new_layout:
        return {}

    result = {}
    for r in range(6, len(grid) + 1):
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

# ============ 項目3：每日訂單快照歷史 ============
# 改成存在 Google Drive（不是本機硬碟），因為 Render 免費方案本機硬碟
# 會在服務休眠/重啟時被清空，存本機的話每次都會把累積的歷史弄丟。

HISTORY_DRIVE_FILENAME = "dashboard_order_history.json"
SEED_HISTORY_FILE = "order_history_seed.json"  # 這個是隨程式碼一起部署的靜態種子檔，本機讀取沒問題


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


DAILY_VALUE_MODELS = {'CR-V(PET)', 'CR-V(e:HEV)'}  # 這幾個欄位是「當天有填=當天有訂單」，不是月累計


def compute_last_order_tracking(history, today_str):
    """掃描累積的每日訂單快照，算出每人每車型「最後一筆訂單」的日期與距今天數。
    兩種欄位語意分開處理：
      - CR-V(PET)/CR-V(e:HEV)：個別配備等級欄位，當天有填數字=當天真的有這個等級的訂單，
        直接找「最近一次有填數字的日期」即可，不需要跟前一天比較。
      - 其他車型：用車型群組的「月累」欄位，是整個月的累計數字，要找數字比前一天增加的那一天。
    """
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
            if model in DAILY_VALUE_MODELS:
                # 當日值欄位：直接找最近一次「當天有填數字」的日期
                for d in dates_sorted:
                    cur_val = history[d].get(p, {}).get(model, 0)
                    if cur_val > 0:
                        last_date = d
            else:
                # 月累計欄位：找數字比前一天增加的那一天(月初重置不算增加，見下方判斷)
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
            result[p][model] = {"last_order_date": last_date, "days_since": days_since, "approx": False}
    return result


MONTH_LAST_DAY_2026 = {
    '1月': (1, 31), '2月': (2, 28), '3月': (3, 31), '4月': (4, 30),
    '5月': (5, 31), '6月': (6, 30), '7月': (7, 31), '8月': (8, 31),
    '9月': (9, 30), '10月': (10, 31), '11月': (11, 30), '12月': (12, 31),
}


def fill_fallback_from_monthly(last_order_tracking, month_sheets_cache, months_to_sum, today_str):
    """快照歷史(order_history)最早只從6/2開始，看不到更早的訂單。
    這裡用跟項目4同一份「月度封存」資料當備援：如果快照歷史完全沒偵測到某人某車型的
    訂單痕跡，就往回找最近一個「該月訂單數>0」的月份，用該月最後一天當近似日期，
    並標記 approx=True（跟先前手動分析的靜態版本邏輯一致）。"""
    today = date.fromisoformat(today_str)
    people_models_with_data = set()
    for m in months_to_sum:
        for p, models in month_sheets_cache.get(m, {}).items():
            for model, vals in models.items():
                if vals.get('訂單', 0) > 0:
                    people_models_with_data.add((p, model))

    for p, model in people_models_with_data:
        if model in DAILY_VALUE_MODELS:
            # CR-V(PET)/CR-V(e:HEV)是「當天有填才算當天有訂單」的欄位，
            # 月度封存表只留得住「封存那一天」的值，退回去用月底當近似值沒有意義
            # （只是剛好告訴你封存那天有沒有訂單，不代表整個月的情況），這裡直接跳過。
            continue
        existing = last_order_tracking.get(p, {}).get(model)
        if existing is not None:
            continue
        # 從最新月份往回找最後一次有訂單的月份
        for m in reversed(months_to_sum):
            v = month_sheets_cache.get(m, {}).get(p, {}).get(model, {}).get('訂單', 0)
            if v > 0:
                mm, dd = MONTH_LAST_DAY_2026[m]
                # 如果剛好是今年最新一個月且今天日期比月底早，用今天日期避免出現「未來日期」
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


TEAM_ORDER = ['林定緯', '林適緯', '陳建道', '陳星佑', '張姉瑀', '歐陽文智', '蔡明憬']

TEAM_STRUCTURE = {
    '歸仁一課': ['林定緯', '林適緯', '陳建道'],
    '歸仁二課': ['陳星佑', '張姉瑀', '歐陽文智', '蔡明憬'],
}
PERSON_TO_DEPT = {p: dept for dept, members in TEAM_STRUCTURE.items() for p in members}
PERSON_TO_DEPT['劉珈微'] = '歸仁一課'  # 已離職，個人列表不顯示，但課別小計仍算她的台數


def compute_dept_totals_scalar(person_value_dict):
    """給 {person: 數字} 的dict，回傳 {'一課': 小計, '二課': 小計}。"""
    totals = {dept: 0 for dept in TEAM_STRUCTURE}
    for p, v in person_value_dict.items():
        dept = PERSON_TO_DEPT.get(p)
        if dept:
            totals[dept] += v
    return totals


def compute_dept_totals_by_model(person_model_dict):
    """給 {person: {model: 數字}} 的dict，回傳 {'一課': {model: 小計}, '二課': {...}}。"""
    totals = {dept: defaultdict(int) for dept in TEAM_STRUCTURE}
    for p, models in person_model_dict.items():
        dept = PERSON_TO_DEPT.get(p)
        if dept:
            for model, v in models.items():
                totals[dept][model] += v
    return {dept: dict(models) for dept, models in totals.items()}


def sort_by_team_order(d):
    """依團隊固定順序(一課→二課)排序dict，不在名單裡的人(如已離職)排在最後。"""
    def key(item):
        name = item[0]
        try:
            return (0, TEAM_ORDER.index(name))
        except ValueError:
            return (1, name)
    return dict(sorted(d.items(), key=key))


def sort_by_total_desc(d):
    """依每人數值加總，由多到少排序。"""
    return dict(sorted(d.items(), key=lambda item: sum(item[1].values()) if isinstance(item[1], dict) else item[1], reverse=True))


# ============ 主流程 ============

def build_daily_snapshot(wb, month_key):
    """組出跟每日正常流程一樣格式的快照：非CR-V車型用月累計(群組層級)，
    CR-V拆成PET/e:HEV用當日值。跟build_dashboard_data()裡的邏輯保持一致，
    這樣不管是正常每日更新還是回溯批次，寫進歷史的資料格式都相同。"""
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


def reset_order_history():
    """清空Drive上存的每日訂單快照歷史(不影響隨程式碼部署的種子檔)。
    修正CR-V欄位驗證邏輯後，之前回溯進去的1~4月錯誤資料需要整個清掉重新回溯，
    這個函式就是做這件事：把Drive上的歷史檔案內容清空成空dict，
    之後重新呼叫 backfill_full_history() 就會是乾淨的重新開始。"""
    from drive_json_store import save_json_to_drive
    save_json_to_drive(DAILY_REPORT_FOLDER_ID, HISTORY_DRIVE_FILENAME, {})
    return {"status": "ok", "message": "Drive上的歷史紀錄已清空，種子檔不受影響"}


def backfill_full_history(max_seconds=100, max_files=30):
    """完整回溯歷史：把Drive資料夾裡所有115年歸仁日報表逐一解析，
    重建出每一天的訂單快照(含CR-V PET/e:HEV當日值)，寫入Drive上的歷史檔案。

    因為檔案數量很多(目前約140+份)，單次執行可能會超過Render的請求逾時限制，
    也可能因為連續開太多份Excel檔案累積記憶體導致OOM，這裡設計成「可分批執行」：
    每次最多處理 max_files 份檔案(用數量硬性限制，比單純看時間更保險)，
    處理到哪裡存到哪裡，已經處理過的日期會跳過，重複呼叫這個函式就能接續進度，
    直到全部處理完。
    """
    import time
    import gc
    start_time = time.time()

    from drive_reader import get_drive_service, download_file
    service = get_drive_service()
    query = f"'{DAILY_REPORT_FOLDER_ID}' in parents and name contains '歸仁日報表115' and trashed = false"
    resp = service.files().list(q=query, pageSize=300, fields="files(id, name)").execute()
    files = resp.get("files", [])

    files_with_date = [(f, _parse_filename_date(f["name"])) for f in files]
    files_with_date = [(f, md) for f, md in files_with_date if md[0] != 0]
    files_with_date.sort(key=lambda x: x[1])  # 由舊到新處理

    history = load_order_history()
    processed = 0
    errors = []
    timed_out = False

    for f, (mm, dd) in files_with_date:
        date_str = f"2026-{mm:02d}-{dd:02d}"
        if date_str in history:
            continue  # 已經處理過，跳過(這樣重複呼叫才能接續進度)
        if processed >= max_files or time.time() - start_time > max_seconds:
            timed_out = True
            break
        try:
            content = download_file(f["id"])
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)
            month_key = f"{mm}月"
            if month_key in wb.sheetnames:
                history[date_str] = build_daily_snapshot(wb, month_key)
                processed += 1
            wb.close()
            del content, wb
            gc.collect()  # 強制回收，避免連續處理多份大檔案累積記憶體導致OOM
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
    team_total_ytd_registration = sum(item1.values())  # 先算總計(含劉珈微)，之後才把她的個人列表拿掉
    item1_dept_totals = compute_dept_totals_scalar(item1)  # 這裡還沒過濾，一課小計會含劉珈微14台

    item4 = {p: {mo: v['領牌'] for mo, v in models.items() if v['領牌'] > 0}
             for p, models in ytd.items()}
    item4 = sort_by_total_desc(item4)

    cur = month_sheets_cache.get(current_month_key, {})
    cur_tracking = month_sheets_cache_for_tracking.get(current_month_key, {})
    month_progress = {}
    for p, models in cur.items():
        month_progress[p] = {
            '訂單': sum(v.get('訂單', 0) for v in models.values()),
            '領牌': sum(v.get('領牌', 0) for v in models.values()),
        }

    # ---- 項目3：最後訂單追蹤（累積每日快照後自動計算，CR-V已拆成PET/e:HEV）----
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
    last_order_tracking = sort_by_team_order(last_order_tracking)

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

            # 手動校正：陳星佑114年1~3月暫時調到永康所，這段期間的領牌算在永康報表裡，
            # 歸仁的114年度報表完全看不到，導致他個人的去年同期基準少算12台
            # （已用永康日報表114_04_22.xls核實：1月4台+2月4台+3月4台=12台）。
            # 這個校正只影響他「個人」的去年同期比較數字，不影響歸仁全所的據點總計。
            KNOWN_CORRECTIONS_LY_YTD = {"陳星佑": 12}
            for _name, _add in KNOWN_CORRECTIONS_LY_YTD.items():
                if _name in ly_item1:
                    ly_item1[_name] += _add
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

    # 劉珈微已離職，個人列表不顯示，但團隊總量(team_total_ytd_registration)已經算過她的數字了
    EXCLUDE_FROM_PERSONAL = {'劉珈微'}
    item1 = {p: v for p, v in item1.items() if p not in EXCLUDE_FROM_PERSONAL}
    item4 = {p: v for p, v in item4.items() if p not in EXCLUDE_FROM_PERSONAL}
    month_progress = {p: v for p, v in month_progress.items() if p not in EXCLUDE_FROM_PERSONAL}
    last_order_tracking = {p: v for p, v in last_order_tracking.items() if p not in EXCLUDE_FROM_PERSONAL}

    # 課別小計（item1_dept_totals 已經在前面算過、含劉珈微，這裡不重算）
    item4_dept_totals = compute_dept_totals_by_model(item4)
    month_progress_dept_totals = compute_dept_totals_by_model(month_progress)
    yoy_last_year_dept_totals = None
    yoy_this_year_dept_totals = None
    if yoy_comparison and yoy_comparison.get("last_year_ytd") is not None:
        # 這裡故意不排除劉珈微：跟item1_dept_totals一樣，課別小計要含她，
        # 只有「個人列表」才不顯示她，兩邊算法要一致
        yoy_last_year_dept_totals = compute_dept_totals_scalar(yoy_comparison["last_year_ytd"])
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
    result = build_dashboard_data()
    print(json.dumps(result, ensure_ascii=False, indent=2))
