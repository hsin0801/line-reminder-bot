"""
歸仁儀表板 - 資料解析模組
負責：從 Google Drive 抓最新的歸仁日報表，解析出儀表板需要的四個項目 + 本月進度，
存成 data.json 供 dashboard_routes.py 讀取渲染。

假設你的 Render 專案已經有 Google 服務帳戶可以存取 Drive（跟你 LINE Bot 抓 W65 用的是同一套）。
如果你現有程式已經有「取得 Drive service」的函式，把 get_drive_service() 換成你自己的即可，
其他解析邏輯不用改。
"""

import os
import io
import json
import base64
from datetime import date, datetime
from collections import defaultdict

import openpyxl
from googleapiclient.discovery import build
from google.oauth2 import service_account

# ============ 設定 ============
DRIVE_FOLDER_ID = os.environ.get("GUEIREN_DRIVE_FOLDER_ID", "1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf")
DATA_FILE = os.environ.get("DASHBOARD_DATA_FILE", "dashboard_data.json")
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

BLACKLIST_SUBSTR = ['合計', '月累', '課', '歸一', '歸二', '歸三', '歸仁', '公司']
MONTHS = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']


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
        'ZRV': 'ZRV', 'CR-V': 'CR-V', 'CRV': 'CR-V', 'HRV': 'HRV', 'HR-V': 'HRV',
        'FIT': 'FIT', 'CIVIC': 'CIVIC', 'ODYSSEY': 'ODYSSEY', 'PRELUDE': 'PRELUDE',
        'ACCORD': 'ACCORD', 'INSIGHT': 'INSIGHT', 'CR-Z': 'CR-Z', 'CRZ': 'CR-Z',
    }
    return mapping.get(n, n)


def find_groups(ws, header_row, sub_row, start_col, end_col_exclusive):
    """回傳每個車型群組 (model_name, start_col, cumcol)。cumcol 取群組內第一個「月累」欄，
    避免抓到後面殘留、無標籤的孤立欄位（這是先前踩過的bug，務必保留 break）。"""
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


# ============ Google Drive ============

def get_drive_service():
    """用服務帳戶憑證建立 Drive API service。
    憑證來源：環境變數 GOOGLE_SERVICE_ACCOUNT_JSON（放整包 service account json 字串，
    或 base64 編碼後的字串都可以，下面會自動判斷）。
    如果你現有程式已經有另一種取得憑證的方式，直接改這個函式即可。"""
    raw = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    try:
        info = json.loads(raw)
    except json.JSONDecodeError:
        info = json.loads(base64.b64decode(raw))
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def find_latest_gueiren_file(service):
    """在指定資料夾裡找檔名含「歸仁日報表」、最新修改時間的檔案"""
    query = (
        f"'{DRIVE_FOLDER_ID}' in parents and "
        f"name contains '歸仁日報表' and trashed = false"
    )
    resp = service.files().list(
        q=query,
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id, name, modifiedTime)",
    ).execute()
    files = resp.get("files", [])
    if not files:
        raise RuntimeError("Drive 資料夾裡找不到歸仁日報表檔案")
    return files[0]


def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    from googleapiclient.http import MediaIoBaseDownload
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


# ============ 核心：組出儀表板資料 ============

def build_dashboard_data():
    service = get_drive_service()
    meta = find_latest_gueiren_file(service)
    buf = download_file(service, meta["id"])
    wb = openpyxl.load_workbook(buf, data_only=True)

    today = date.today()
    current_month_key = f"{today.month}月"

    # ---- item1 + item4: 累計所有已結束的月份 sheet ----
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

    # ---- 本月進度 ----
    month_progress = {}
    if current_month_key in wb.sheetnames:
        cur = parse_month_sheet(wb, current_month_key)
        for p, models in cur.items():
            month_progress[p] = {
                '訂單': sum(v.get('訂單', 0) for v in models.values()),
                '領牌': sum(v.get('領牌', 0) for v in models.values()),
            }

    data = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "source_file": meta["name"],
        "item1_ytd_registration": item1,
        "item4_ytd_by_model": item4,
        "month_progress": month_progress,
        # item2(去年同期比較) 和 item3(最後訂單追蹤) 需要額外的歷史檔案/去年檔案，
        # 建議另外寫一支週期性任務把結果算好後併進這個 json，
        # 或先手動維護，avoid每次重跑都要重新下載大量歷史檔案。
    }

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return data


if __name__ == "__main__":
    result = build_dashboard_data()
    print(json.dumps(result, ensure_ascii=False, indent=2))
