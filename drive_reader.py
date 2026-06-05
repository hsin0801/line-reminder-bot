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
    date_match = re.search(r'(\d{8})', file_name)
    password = date_match.group(1) if date_match else target_date.strftime("%Y%m%d")

    content = download_file(file_info["id"])
    decrypted = decrypt_xls(content, password)
    wb = xlrd.open_workbook(file_contents=decrypted)
    sh = wb.sheet_by_index(0)

    guiren_col = None
    yongkang_col = None
    for r in range(sh.nrows):
        for c in range(sh.ncols):
            v = str(sh.cell_value(r, c)).strip()
            if v == "歸仁" and guiren_col is None:
                guiren_col = c
            if v == "永康" and yongkang_col is None:
                yongkang_col = c
        if guiren_col and yongkang_col:
            break

    if not guiren_col:
        raise Exception(f"找不到歸仁欄位，檔案: {file_name}")

    result = {"file": file_name, "date": password, "歸仁": {}, "永康": {}}
    cur_model = ""
    for r in range(sh.nrows):
        c0 = str(sh.cell_value(r, 0)).strip()
        c1 = str(sh.cell_value(r, 1)).strip()
        if c0 and c0 not in ("", "0430", "0603", "0604"):
            cur_model = c0
        if c1 in ("Booking", "Register", "目標達成率", "進度達成率", "來店"):
            label = f"{cur_model}_{c1}" if cur_model else c1
            gval = sh.cell_value(r, guiren_col)
            yval = sh.cell_value(r, yongkang_col) if yongkang_col else "-"
            if c1 in ("目標達成率", "進度達成率"):
                if isinstance(gval, float) and gval > 0:
                    gval = f"{gval*100:.1f}%"
                if isinstance(yval, float) and yval > 0:
                    yval = f"{yval*100:.1f}%"
            if gval not in (0, 0.0, "", None):
                result["歸仁"][label] = gval
            if yval not in (0, 0.0, "", None):
                result["永康"][label] = yval
    return result

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
