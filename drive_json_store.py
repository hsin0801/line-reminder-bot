"""
Drive JSON 持久化工具 - 給歸仁/永康共用
------------------------------------------------
背景：Render 免費方案的本機硬碟是「暫時性」的——服務閒置一段時間會自動休眠，
休眠後醒來(或每次重新部署)本機檔案系統會被清空重來，之前存在本機的
「每日訂單快照歷史」(dashboard_order_history.json / yongkang_order_history.json)
就會整個消失，項目3(最後一筆訂單追蹤)的準確度也會跟著被打回原點。

解法：改把這份會持續累積的小檔案存進 Google Drive 裡的一個檔案，
每次讀取/寫入都直接對 Drive 操作，不再依賴本機硬碟。Drive 上的檔案
不會因為 Render 服務重啟而消失，資料就能真的持續累積下去。
"""

import io
import os
import json

# 注意：drive_reader.py 裡的 get_drive_service() 用的是唯讀權限
# (SCOPES = drive.readonly)，拿來寫入(create/update)一定會被拒絕，
# 就算服務帳戶本身在Drive上有編輯權限也一樣——因為OAuth核發的權杖
# 範圍(scope)本身就鎖死唯讀，不是共用權限的問題。
# 這裡另外用同一組帳密憑證(GOOGLE_CREDENTIALS_JSON)但要求可寫入的scope，
# 建立一個獨立的、有寫入權限的Drive service。
_WRITE_SCOPES = ["https://www.googleapis.com/auth/drive"]


def _get_writable_drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    creds_dict = json.loads(os.environ.get("GOOGLE_CREDENTIALS_JSON", "{}"))
    creds = Credentials.from_service_account_info(creds_dict, scopes=_WRITE_SCOPES)
    return build("drive", "v3", credentials=creds)


def _find_file_in_folder(service, folder_id, filename):
    query = f"'{folder_id}' in parents and name = '{filename}' and trashed = false"
    resp = service.files().list(q=query, pageSize=5, fields="files(id, name)").execute()
    files = resp.get("files", [])
    return files[0] if files else None


def load_json_from_drive(folder_id, filename):
    """從 Drive 資料夾裡讀取指定檔名的 JSON 內容，不存在就回傳空 dict。
    讀取用唯讀service即可(沿用drive_reader既有的download_file)。"""
    from drive_reader import get_drive_service, download_file

    service = get_drive_service()
    existing = _find_file_in_folder(service, folder_id, filename)
    if not existing:
        return {}
    content = download_file(existing["id"])
    try:
        return json.loads(content.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return {}


def save_json_to_drive(folder_id, filename, data):
    """把 JSON 內容存進 Drive 資料夾裡的指定檔名，檔案已存在就更新內容，
    不存在就新建一個。這裡一定要用可寫入權限的service，唯讀service會被拒絕。"""
    from googleapiclient.http import MediaIoBaseUpload

    service = _get_writable_drive_service()
    content_bytes = json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(content_bytes), mimetype="application/json", resumable=False)

    existing = _find_file_in_folder(service, folder_id, filename)
    if existing:
        service.files().update(fileId=existing["id"], media_body=media).execute()
    else:
        file_metadata = {"name": filename, "parents": [folder_id]}
        service.files().create(body=file_metadata, media_body=media, fields="id").execute()
