"""
guiren_kpi_checker.py
每天日報表更新後，比對日報表 vs 週邊指標的配件/保費金額
有差異時發 LINE 通知給劉宗鑫
"""

import io, logging, zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from collections import defaultdict
import pytz

TZ       = pytz.timezone('Asia/Taipei')
LINE_UID = 'U272a3c6b1f3d10a3677769cb4f73fe1d'  # 劉宗鑫 LINE ID

KPI_FILE_ID  = '1-8JmC6MlF-WPMsgeWqh43f_q5BiA3Q_G'  # 週邊指標
DAILY_FOLDER = '1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf'  # 日報表資料夾
COMPANY      = '巧克力汽車商業股份有限公司'

SA_ALL = ['林定緯','林適緯','劉珈微','陳建道','陳星佑','蔡明憬','張姉瑀','歐陽文智']
SA_MAP = {'定緯':'林定緯','適緯':'林適緯','珈微':'劉珈微','建道':'陳建道',
          '星佑':'陳星佑','姉瑀':'張姉瑀','文智':'歐陽文智','明憬':'蔡明憬'}

ACC_DIFF_THRESHOLD  = 500   # 配件差異門檻（元）
PREM_DIFF_THRESHOLD = 500   # 保費差異門檻（元）


# ── 共用：下載 + XML 解析 ──────────────────────────────────
def _download(service, file_id):
    return service.files().get_media(fileId=file_id).execute()

def _latest_daily(service):
    result = service.files().list(
        q=f"'{DAILY_FOLDER}' in parents and name contains '歸仁日報表' "
          "and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
        orderBy='modifiedTime desc', pageSize=1,
        fields='files(id,name)'
    ).execute()
    items = result.get('files', [])
    return (items[0]['id'], items[0]['name']) if items else (None, None)

def _parse_sheets(xlsx_bytes, target_sheets):
    """zipfile + XML 解析，只讀目標 sheet，低記憶體"""
    import re
    result = {}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
        names = zf.namelist()

        # sharedStrings
        shared = []
        if 'xl/sharedStrings.xml' in names:
            with zf.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in tree.findall('.//ns:si', ns):
                    t_nodes = si.findall('.//ns:t', ns)
                    shared.append(''.join(t.text or '' for t in t_nodes))

        # sheet name → rid
        sheet_map = {}
        if 'xl/workbook.xml' in names:
            with zf.open('xl/workbook.xml') as f:
                tree = ET.parse(f)
                for sh in tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                    sname = sh.get('name', '')
                    rid   = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
                    sheet_map[sname] = rid

        # rid → path
        rid_to_path = {}
        rels = 'xl/_rels/workbook.xml.rels'
        if rels in names:
            with zf.open(rels) as f:
                tree = ET.parse(f)
                for rel in tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rid_to_path[rel.get('Id','')] = 'xl/' + rel.get('Target','').lstrip('/')

        def parse_cell(c):
            t = c.get('t', '')
            vn = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if vn is None: return None
            val = vn.text or ''
            if t == 's':
                try: return shared[int(val)]
                except: return val
            try: return float(val) if '.' in val else int(val)
            except: return val

        def col_idx(ref):
            m = re.match(r'([A-Z]+)', ref or '')
            if not m: return 0
            s = m.group(1); r = 0
            for ch in s: r = r * 26 + (ord(ch) - 64)
            return r - 1

        for sname in target_sheets:
            if sname not in sheet_map: continue
            path = rid_to_path.get(sheet_map[sname])
            if not path or path not in names: continue
            rows_data = []
            with zf.open(path) as f:
                tree = ET.parse(f)
                ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                for row_el in tree.findall(f'.//{{{ns}}}row'):
                    cells = row_el.findall(f'{{{ns}}}c')
                    if not cells: rows_data.append([]); continue
                    max_col = col_idx(cells[-1].get('r', ''))
                    row = [None] * (max_col + 1)
                    for c in cells:
                        ci = col_idx(c.get('r', ''))
                        row[ci] = parse_cell(c)
                    rows_data.append(row)
            result[sname] = rows_data
    return result


# ── 1. 從日報表讀當月累計 ─────────────────────────────────
def read_daily_curr(service):
    """
    回傳 { sa: { reg, acc, prem } }
    reg  = 月累領牌台數  (col31, idx30)
    acc  = 配件月累總額  (col162, idx161)
    prem = 保費月累      (col44, idx43)
    """
    now = datetime.now(TZ)
    curr_month = now.month
    file_id, _ = _latest_daily(service)
    if not file_id:
        return {}, curr_month

    raw = _download(service, file_id)
    sheets = _parse_sheets(raw, [f'{curr_month}月'])
    del raw

    ws = sheets.get(f'{curr_month}月', [])
    result = {}
    for row in ws:
        if not row: continue
        name = str(row[0] or '').strip()
        if name not in SA_ALL: continue
        def g(i): return int(row[i]) if i < len(row) and isinstance(row[i], (int, float)) else 0
        result[name] = {
            'reg':  g(30),   # col31 月累領牌
            'acc':  g(161),  # col162 配件月累總額
            'prem': g(43),   # col44 保費月累
        }
    return result, curr_month


# ── 2. 從週邊指標讀當月（已領牌筆數+金額）────────────────
def read_kpi_curr(service, curr_month):
    """
    只計算已領牌（非租車、非公司車）的台數、配件加總、保費加總
    回傳 { sa: { reg, acc, prem } }
    """
    raw = _download(service, KPI_FILE_ID)
    sname = f'2026.{curr_month:02d}月'
    sheets = _parse_sheets(raw, [sname])
    del raw

    result = defaultdict(lambda: {'reg': 0, 'acc': 0, 'prem': 0})
    for row in sheets.get(sname, []):
        if not row or len(row) < 7: continue
        sa_abbr  = str(row[1] or '').strip()
        customer = str(row[2] or '').strip()
        acc      = row[3]
        prem     = row[7] if len(row) > 7 else None
        if sa_abbr not in SA_MAP: continue
        sa = SA_MAP[sa_abbr]
        if COMPANY in customer: continue
        if customer == '租車': continue  # 租車不算領牌件數
        result[sa]['reg']  += 1
        if isinstance(acc,  (int, float)): result[sa]['acc']  += acc
        if isinstance(prem, (int, float)): result[sa]['prem'] += prem

    return dict(result)


# ── 3. 比對 + 產生通知訊息 ───────────────────────────────
def check_diff(daily, kpi):
    """
    比對每個 SA：
    1. 台數不一致 → 警告（週邊指標有未領牌筆數）
    2. 台數一致但配件差 > 500 → 警告
    3. 台數一致但保費差 > 500 → 警告
    回傳 issues list
    """
    issues = []
    for sa in SA_ALL:
        d = daily.get(sa, {'reg': 0, 'acc': 0, 'prem': 0})
        k = kpi.get(sa,   {'reg': 0, 'acc': 0, 'prem': 0})

        d_reg = d['reg'];  k_reg = k['reg']
        d_acc = d['acc'];  k_acc = k['acc']
        d_prm = d['prem']; k_prm = k['prem']

        sa_issues = []

        if k_reg != d_reg:
            sa_issues.append(
                f"  ・台數不符：日報表 {d_reg} 台 / 週邊指標 {k_reg} 台"
                + ("（週邊多了未領牌筆數）" if k_reg > d_reg else "（週邊少了已領牌筆數）")
            )

        if d_reg == k_reg and d_reg > 0:
            acc_diff = d_acc - k_acc
            if abs(acc_diff) >= ACC_DIFF_THRESHOLD:
                sa_issues.append(
                    f"  ・配件金額差異：日報表 {d_acc:,} / 週邊指標 {k_acc:,}"
                    f"（差 {acc_diff:+,} 元）"
                )
            prm_diff = d_prm - k_prm
            if abs(prm_diff) >= PREM_DIFF_THRESHOLD:
                sa_issues.append(
                    f"  ・保費金額差異：日報表 {d_prm:,} / 週邊指標 {k_prm:,}"
                    f"（差 {prm_diff:+,} 元）"
                )

        if sa_issues:
            issues.append(f"【{sa}】\n" + "\n".join(sa_issues))

    return issues


# ── 4. 發 LINE 通知 ───────────────────────────────────────
def _send_line(message):
    """發 LINE push message"""
    import requests, os
    token = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN', '')
    if not token:
        logging.error('[KPI_CHECK] LINE token 未設定')
        return
    requests.post(
        'https://api.line.me/v2/bot/message/push',
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {token}'
        },
        json={
            'to': LINE_UID,
            'messages': [{'type': 'text', 'text': message}]
        },
        timeout=10
    )


# ── 主入口 ────────────────────────────────────────────────
def run_kpi_check(service):
    """
    從 app.py 的 cron job 呼叫：
      from guiren_kpi_checker import run_kpi_check
      run_kpi_check(get_drive_service())
    """
    now = datetime.now(TZ)
    try:
        daily, curr_month = read_daily_curr(service)
        kpi   = read_kpi_curr(service, curr_month)
        issues = check_diff(daily, kpi)

        if not issues:
            logging.info('[KPI_CHECK] 本日比對無差異')
            return

        month_name = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'][curr_month]
        header = (
            f"⚠️ 歸仁 {month_name}月 日報表 vs 週邊指標 差異提醒\n"
            f"（{now.strftime('%m/%d %H:%M')} 更新）\n"
            f"{'─'*22}\n"
        )
        body = "\n\n".join(issues)
        _send_line(header + body)
        logging.info(f'[KPI_CHECK] 發現 {len(issues)} 筆差異，已通知')

    except Exception as e:
        logging.error(f'[KPI_CHECK] 執行失敗：{e}')


if __name__ == '__main__':
    print('guiren_kpi_checker.py loaded OK')
