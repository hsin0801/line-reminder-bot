import os
import io
import json
import re
from datetime import date

SPEED_REPORT_FOLDER_ID = os.environ.get("SPEED_REPORT_FOLDER_ID", "1zUKKS0G1vsbnbptJLtU2HYTFGdIt8bv7")
DAILY_REPORT_FOLDER_ID = os.environ.get("DAILY_REPORT_FOLDER_ID", "1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf")
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

def get_drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    creds_dict = json.loads(os.environ.get("GOOGLE_CREDENTIALS_JSON", "{}"))
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)

def get_latest_file(folder_id, keyword=""):
    service = get_drive_service()
    query = f"'{folder_id}' in parents and trashed=false"
    if keyword:
        query += f" and name contains '{keyword}'"
    results = service.files().list(
        q=query, orderBy="createdTime desc", pageSize=5,
        fields="files(id, name, createdTime)"
    ).execute()
    files = results.get("files", [])
    return files[0] if files else None

def download_file(file_id):
    from googleapiclient.http import MediaIoBaseDownload
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()

def decrypt_xls(content, password):
    import msoffcrypto
    buf_in = io.BytesIO(content)
    office = msoffcrypto.OfficeFile(buf_in)
    if office.is_encrypted():
        office.load_key(password=password)
        buf_out = io.BytesIO()
        office.decrypt(buf_out)
        return buf_out.getvalue()
    return content

def find_store_cols(sh):
    """
    找月累欄位邏輯：
    - 找到據點名稱（如「歸仁」）所在欄 c
    - 下一列的 c-1 欄必須是「月累」
    - 符合條件才採用 c-1 作為該據點月累欄
    - 同一列必須同時找到歸仁＋永康才算有效
    """
    targets = {"歸仁", "永康", "東台南", "西台南"}
    for r in range(sh.nrows):
        found = {}
        for c in range(1, sh.ncols):
            v = str(sh.cell_value(r, c)).strip()
            if v in targets and v not in found:
                # 驗證：下一列的 c-1 欄是否為「月累」
                if r + 1 < sh.nrows:
                    check = str(sh.cell_value(r + 1, c - 1)).strip()
                    if check == "月累":
                        found[v] = c - 1  # 月累欄
        if "歸仁" in found and "永康" in found:
            return found, r
    return None, None

def get_speed_report(target_date=None):
    import xlrd
    if target_date is None:
        target_date = date.today()

    file_info = get_latest_file(SPEED_REPORT_FOLDER_ID, "業績速報")
    if not file_info:
        raise Exception(f"找不到業績速報，資料夾ID: {SPEED_REPORT_FOLDER_ID}")

    file_name = file_info["name"]
    all_dates = re.findall(r'\d{8}', file_name)
    password = all_dates[-1] if all_dates else target_date.strftime("%Y%m%d")

    content = download_file(file_info["id"])
    decrypted = decrypt_xls(content, password)
    wb = xlrd.open_workbook(file_contents=decrypted)
    sh = wb.sheet_by_index(0)

    cols, header_row = find_store_cols(sh)
    if cols is None:
        raise Exception(f"找不到據點欄位，檔案: {file_name}")

    result = {
        "file": file_name, "date": password,
        "歸仁": {}, "永康": {}, "東台南": {}, "西台南": {}
    }

    cur_model = ""
    for r in range(header_row, sh.nrows):
        c0 = str(sh.cell_value(r, 0)).strip()
        c1 = str(sh.cell_value(r, 1)).strip()
        if c0 and c0 not in ("", "0430", "0603", "0604"):
            cur_model = c0
        if c1 in ("Booking", "Register", "目標達成率", "進度達成率", "來店"):
            label = f"{cur_model}_{c1}" if cur_model else c1
            for name, col in cols.items():
                if col is None:
                    continue
                # 已存在的 key 不覆蓋（保留第一次出現的值）
                if label in result[name]:
                    continue
                val = sh.cell_value(r, col)
                if c1 in ("目標達成率", "進度達成率"):
                    if isinstance(val, float) and val > 0:
                        val = f"{val*100:.1f}%"
                if val not in (0, 0.0, "", None):
                    result[name][label] = val
    return result


def format_speed_report_message(report):
    date_str = report.get("date", "")
    display_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:]}" if len(date_str) == 8 else date_str

    def get_first_val(store, suffix):
        # 優先取 TOTAL 合計列，再 fallback 到第一個符合的
        total_key = "TOTAL" + suffix
        if total_key in report[store]:
            try:
                return int(float(report[store][total_key]))
            except:
                pass
        for k, v in report[store].items():
            if k.endswith(suffix):
                try:
                    return int(float(v))
                except:
                    pass
        return 0

    def get_first_pct(store, suffix):
        total_key = "TOTAL" + suffix
        if total_key in report[store] and isinstance(report[store][total_key], str) and "%" in report[store][total_key]:
            return report[store][total_key]
        for k, v in report[store].items():
            if k.endswith(suffix) and isinstance(v, str) and "%" in v:
                return v
        return "-"

    gj_bk  = get_first_val("歸仁",  "_Booking")
    yk_bk  = get_first_val("永康",  "_Booking")
    et_bk  = get_first_val("東台南", "_Booking")
    wt_bk  = get_first_val("西台南", "_Booking")
    gj_rg  = get_first_val("歸仁",  "_Register")
    yk_rg  = get_first_val("永康",  "_Register")
    et_rg  = get_first_val("東台南", "_Register")
    wt_rg  = get_first_val("西台南", "_Register")
    gj_ach = get_first_pct("歸仁",  "_目標達成率")
    yk_ach = get_first_pct("永康",  "_目標達成率")

    choc_bk = gj_bk + yk_bk
    sy_bk   = et_bk + wt_bk
    choc_rg = gj_rg + yk_rg
    sy_rg   = et_rg + wt_rg
    bk_diff = choc_bk - sy_bk
    diff_symbol = "▲" if bk_diff >= 0 else "▼"

    msg = f"""📊 {display_date} 業績速報

━━━━━━━━━━━━━━
🍫 巧克力（訂單月累）
  歸仁：{gj_bk} 台　永康：{yk_bk} 台
  合計：{choc_bk} 台
  歸仁達成率：{gj_ach}　永康：{yk_ach}

⚔️ 伸陽（訂單月累）
  東台南：{et_bk} 台　西台南：{wt_bk} 台
  合計：{sy_bk} 台

📌 巧克力 vs 伸陽：{diff_symbol}{abs(bk_diff)} 台
━━━━━━━━━━━━━━
領牌月累
  巧克力：{choc_rg} 台（歸仁 {gj_rg}／永康 {yk_rg}）
  伸陽：{sy_rg} 台
━━━━━━━━━━━━━━"""
    return msg.strip()


def get_daily_report(target_date=None):
    import xlrd
    if target_date is None:
        target_date = date.today()

    file_info = get_latest_file(DAILY_REPORT_FOLDER_ID, "歸仁日報表")
    if not file_info:
        file_info = get_latest_file(DAILY_REPORT_FOLDER_ID, "日報表")
    if not file_info:
        raise Exception(f"找不到日報表，資料夾ID: {DAILY_REPORT_FOLDER_ID}")

    file_name = file_info["name"]
    content = download_file(file_info["id"])
    members = ["林定緯", "林適緯", "陳建道", "陳星佑", "張姉瑀", "歐陽文智", "蔡明憬"]
    result = {"file": file_name, "data": {}}

    if file_name.endswith(".xlsx"):
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)
        today = date.today()
        ws = None
        for name in wb.sheetnames:
            if str(today.month) in name:
                ws = wb[name]
                break
        if not ws:
            ws = wb.active
        for row in ws.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                name = str(cell).strip() if cell else ""
                if name in members:
                    row_vals = [v for v in row if v is not None]
                    result["data"][name] = str(row_vals[:15])
                    break
        wb.close()
    else:
        wb = xlrd.open_workbook(file_contents=content)
        sh = wb.sheet_by_index(0)
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = str(sh.cell_value(r, c)).strip()
                if v in members and v not in result["data"]:
                    row_data = [sh.cell_value(r, cc) for cc in range(min(sh.ncols, 15))]
                    result["data"][v] = str(row_data)
    return result
