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
        q=query,
        orderBy="createdTime desc",
        pageSize=5,
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

def get_speed_report(target_date=None):
    import xlrd
    if target_date is None:
        target_date = date.today()

    file_info = get_latest_file(SPEED_REPORT_FOLDER_ID, "業績速報")
    if not file_info:
        raise Exception(f"找不到業績速報，資料夾ID: {SPEED_REPORT_FOLDER_ID}")

    file_name = file_info["name"]
    # 密碼規則：取檔名最後一組日期（20260606-0607 → 20260607）
    all_dates = re.findall(r'\d{8}', file_name)
    if all_dates:
        password = all_dates[-1]
    else:
        password = target_date.strftime("%Y%m%d")

    content = download_file(file_info["id"])
    decrypted = decrypt_xls(content, password)
    wb = xlrd.open_workbook(file_contents=decrypted)
    sh = wb.sheet_by_index(0)

    # 找四個據點的欄位
    cols = {"歸仁": None, "永康": None, "東台南": None, "西台南": None}
    for r in range(sh.nrows):
        for c in range(sh.ncols):
            v = str(sh.cell_value(r, c)).strip()
            if v in cols and cols[v] is None:
                cols[v] = c
        if all(v is not None for v in cols.values()):
            break

    if cols["歸仁"] is None:
        raise Exception(f"找不到歸仁欄位，檔案: {file_name}")

    result = {
        "file": file_name,
        "date": password,
        "歸仁": {},
        "永康": {},
        "東台南": {},
        "西台南": {}
    }

    cur_model = ""
    for r in range(sh.nrows):
        c0 = str(sh.cell_value(r, 0)).strip()
        c1 = str(sh.cell_value(r, 1)).strip()
        if c0 and c0 not in ("", "0430", "0603", "0604"):
            cur_model = c0
        if c1 in ("Booking", "Register", "目標達成率", "進度達成率", "來店"):
            label = f"{cur_model}_{c1}" if cur_model else c1
            for name, col in cols.items():
                if col is None:
                    continue
                val = sh.cell_value(r, col)
                if c1 in ("目標達成率", "進度達成率"):
                    if isinstance(val, float) and val > 0:
                        val = f"{val*100:.1f}%"
                if val not in (0, 0.0, "", None):
                    result[name][label] = val

    return result


def format_speed_report_message(report):
    """
    將 get_speed_report() 的結果整理成 LINE 推播訊息文字。
    巧克力（歸仁＋永康）合計 vs 伸陽（東台南＋西台南）合計。
    """
    date_str = report.get("date", "")
    display_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:]}" if len(date_str) == 8 else date_str

    def get_val(store, key_contains, field):
        """從指定據點取特定車款+欄位的數字，key_contains='' 取 TOTAL"""
        for k, v in report[store].items():
            if key_contains in k and k.endswith(f"_{field}"):
                try:
                    return float(v) if not isinstance(v, str) else v
                except:
                    return v
        return 0

    def total_booking(stores):
        """合計多個據點的月累 Booking"""
        total = 0
        for s in stores:
            for k, v in report[s].items():
                if k == "TOTAL_Booking" or (k.endswith("_Booking") and "TOTAL" in k):
                    try:
                        total += float(v)
                    except:
                        pass
                    break
            else:
                # fallback：找第一個含 Booking 且不含車款的
                for k, v in report[s].items():
                    if k.endswith("_Booking"):
                        try:
                            total += float(v)
                        except:
                            pass
                        break
        return int(total)

    def total_register(stores):
        total = 0
        for s in stores:
            for k, v in report[s].items():
                if k.endswith("_Register"):
                    try:
                        total += float(v)
                    except:
                        pass
                    break
        return int(total)

    def get_booking_mtd(store):
        """取月累 Booking（第一個 Booking 欄位）"""
        for k, v in report[store].items():
            if k.endswith("_Booking"):
                try:
                    return int(float(v))
                except:
                    pass
        return 0

    def get_register_mtd(store):
        for k, v in report[store].items():
            if k.endswith("_Register"):
                try:
                    return int(float(v))
                except:
                    pass
        return 0

    def get_achieve(store, field):
        for k, v in report[store].items():
            if k.endswith(f"_{field}") and isinstance(v, str) and "%" in v:
                return v
        return "-"

    # 巧克力合計
    choc_bk = get_booking_mtd("歸仁") + get_booking_mtd("永康")
    choc_rg = get_register_mtd("歸仁") + get_register_mtd("永康")
    gj_bk   = get_booking_mtd("歸仁")
    yk_bk   = get_booking_mtd("永康")
    gj_rg   = get_register_mtd("歸仁")
    yk_rg   = get_register_mtd("永康")

    # 伸陽合計
    sy_bk = get_booking_mtd("東台南") + get_booking_mtd("西台南")
    sy_rg = get_register_mtd("東台南") + get_register_mtd("西台南")
    et_bk = get_booking_mtd("東台南")
    wt_bk = get_booking_mtd("西台南")

    # 達成率（取歸仁為代表，或可分開顯示）
    gj_bk_ach = get_achieve("歸仁", "目標達成率")
    yk_bk_ach = get_achieve("永康", "目標達成率")

    # 差距
    bk_diff = choc_bk - sy_bk
    diff_symbol = "▲" if bk_diff >= 0 else "▼"
    diff_abs = abs(bk_diff)

    msg = f"""📊 {display_date} 業績速報

━━━━━━━━━━━━━━
🍫 巧克力（訂單月累）
  歸仁：{gj_bk} 台　永康：{yk_bk} 台
  合計：{choc_bk} 台
  歸仁達成率：{gj_bk_ach}　永康：{yk_bk_ach}

⚔️ 伸陽（訂單月累）
  東台南：{et_bk} 台　西台南：{wt_bk} 台
  合計：{sy_bk} 台

📌 巧克力 vs 伸陽：{diff_symbol}{diff_abs} 台
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
