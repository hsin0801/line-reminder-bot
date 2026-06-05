import os
import io
import json
import base64
import msoffcrypto
import xlrd
import openpyxl
from datetime import date, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SPEED_REPORT_FOLDER_ID = os.environ.get("SPEED_REPORT_FOLDER_ID", "1zUKKS0G1vsbnbptJLtU2HYTFGdIt8bv7")
DAILY_REPORT_FOLDER_ID = os.environ.get("DAILY_REPORT_FOLDER_ID", "1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly"
]

def get_drive_service():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise Exception("缺少 GOOGLE_CREDENTIALS_JSON")
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)

def get_latest_file(folder_id: str, keyword: str = "") -> dict:
    """取得資料夾內最新的檔案"""
    service = get_drive_service()
    query = f"'{folder_id}' in parents and trashed=false"
    if keyword:
        query += f" and name contains '{keyword}'"
    results = service.files().list(
        q=query,
        orderBy="createdTime desc",
        pageSize=5,
        fields="files(id, name, createdTime, mimeType)"
    ).execute()
    files = results.get("files", [])
    if not files:
        return None
    return files[0]

def download_file(file_id: str) -> bytes:
    """下載檔案內容"""
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()

def decrypt_xls(content: bytes, password: str) -> bytes:
    """解密加密的 xls 檔案"""
    try:
        buf_in = io.BytesIO(content)
        office = msoffcrypto.OfficeFile(buf_in)
        if office.is_encrypted():
            office.load_key(password=password)
            buf_out = io.BytesIO()
            office.decrypt(buf_out)
            return buf_out.getvalue()
        return content
    except Exception as e:
        print(f"[ERROR] 解密失敗: {e}")
        return content

# ─────────────────────────────────────────
#  業績速報分析
# ─────────────────────────────────────────
def get_speed_report(target_date: date = None) -> dict:
    """
    讀取最新業績速報，回傳歸仁、永康等重點數據
    target_date: 指定日期（預設今天）
    """
    if target_date is None:
        target_date = date.today()

    password = target_date.strftime("%Y%m%d")

    try:
        # 找最新的業績速報
        file_info = get_latest_file(SPEED_REPORT_FOLDER_ID, "業績速報")
        if not file_info:
            return {"error": "找不到業績速報檔案"}

        file_name = file_info["name"]
        print(f"[速報] 讀取: {file_name}")

        # 從檔名猜密碼（格式：DLR業績速報20260603.xls）
        import re
        date_match = re.search(r'(\d{8})', file_name)
        if date_match:
            password = date_match.group(1)

        # 下載並解密
        content = download_file(file_info["id"])
        decrypted = decrypt_xls(content, password)

        # 讀取 Excel
        wb = xlrd.open_workbook(file_contents=decrypted)
        sh = wb.sheet_by_index(0)

        # 找歸仁欄位（row 79 是據點標題列）
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
            return {"error": "找不到歸仁欄位", "file": file_name}

        # 讀取重要指標
        result = {
            "file": file_name,
            "date": password,
            "歸仁": {},
            "永康": {},
        }

        # 找 Booking、Register、目標達成率 等行
        cur_model = ""
        for r in range(sh.nrows):
            c0 = str(sh.cell_value(r, 0)).strip()
            c1 = str(sh.cell_value(r, 1)).strip()
            if c0 and c0 not in ("", "0430"):
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

    except Exception as e:
        return {"error": f"讀取業績速報失敗: {str(e)}"}


# ─────────────────────────────────────────
#  日報表分析
# ─────────────────────────────────────────
def get_daily_report(target_date: date = None) -> dict:
    """
    讀取最新歸仁日報表，回傳今日業績摘要
    """
    if target_date is None:
        target_date = date.today()

    try:
        # 找歸仁日報表（排除永康）
        file_info = get_latest_file(DAILY_REPORT_FOLDER_ID, "歸仁日報表")
        if not file_info:
            # 備援：找任何日報表
            file_info = get_latest_file(DAILY_REPORT_FOLDER_ID, "日報表")
        if not file_info:
            return {"error": "找不到日報表檔案"}

        file_name = file_info["name"]
        print(f"[日報] 讀取: {file_name}")

        content = download_file(file_info["id"])

        # 判斷格式
        if file_name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        else:
            # .xls 格式
            wb_xls = xlrd.open_workbook(file_contents=content)
            # 轉成可讀格式
            return _parse_xls_daily(wb_xls, file_name)

        # 解析 xlsx（6月格式）
        return _parse_xlsx_daily(wb, file_name)

    except Exception as e:
        return {"error": f"讀取日報表失敗: {str(e)}"}


def _parse_xlsx_daily(wb, file_name: str) -> dict:
    """解析 xlsx 格式的日報表"""
    result = {
        "file": file_name,
        "一課": {},
        "二課": {},
        "全體": {}
    }
    try:
        # 找當月工作表
        today = date.today()
        month_str = str(today.month)
        ws = None
        for name in wb.sheetnames:
            if month_str in name or "6" in name:
                ws = wb[name]
                break
        if not ws:
            ws = wb.active

        # 讀取數據：找各業務員的訂單/領牌數
        members = {
            "林定緯": None, "林適緯": None, "陳建道": None,
            "陳星佑": None, "張姉瑀": None, "歐陽文智": None, "蔡明憬": None
        }

        for row in ws.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                name = str(cell).strip() if cell else ""
                if name in members:
                    # 讀取同行的訂單、領牌數
                    row_vals = [str(v) if v is not None else "" for v in row]
                    members[name] = row_vals
                    break

        result["members"] = members
        return result
    except Exception as e:
        result["error"] = str(e)
        return result


def _parse_xls_daily(wb, file_name: str) -> dict:
    """解析 xls 格式的日報表"""
    result = {"file": file_name, "data": {}}
    try:
        sh = wb.sheet_by_index(0)
        members = ["林定緯", "林適緯", "陳建道", "陳星佑", "張姉瑀", "歐陽文智", "蔡明憬"]

        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = str(sh.cell_value(r, c)).strip()
                if v in members:
                    row_data = []
                    for cc in range(min(sh.ncols, 30)):
                        row_data.append(sh.cell_value(r, cc))
                    result["data"][v] = row_data
        return result
    except Exception as e:
        result["error"] = str(e)
        return result


# ─────────────────────────────────────────
#  格式化輸出（給 LINE 訊息用）
# ─────────────────────────────────────────
def format_speed_report_summary(data: dict) -> str:
    """把業績速報格式化成 LINE 訊息"""
    if "error" in data:
        return f"❌ 業績速報讀取失敗：{data['error']}"

    lines = [f"⚡ 業績速報 {data.get('date', '')}"]
    lines.append("=" * 20)

    for dealer in ["歸仁", "永康"]:
        if dealer not in data or not data[dealer]:
            continue
        lines.append(f"\n📍 {dealer}")
        for key, val in data[dealer].items():
            if isinstance(val, float) and val == int(val):
                val = int(val)
            lines.append(f"  {key}: {val}")

    return "\n".join(lines)


def format_daily_report_summary(data: dict) -> str:
    """把日報表格式化成 LINE 訊息"""
    if "error" in data:
        return f"❌ 日報表讀取失敗：{data['error']}"

    lines = [f"📋 歸仁日報表 {data.get('file', '')}"]
    lines.append("=" * 20)

    if "members" in data:
        for name, vals in data["members"].items():
            if vals:
                lines.append(f"  {name}: {vals}")

    return "\n".join(lines)
