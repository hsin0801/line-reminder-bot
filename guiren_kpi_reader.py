"""
guiren_kpi_reader.py v4
低記憶體版：用 zipfile+xml 直接解析，不用 openpyxl
"""
import io, json, logging, zipfile
from datetime import datetime
from collections import defaultdict
import pytz

TZ = pytz.timezone('Asia/Taipei')

SA_DISPLAY = ['林定緯','林適緯','陳建道','陳星佑','張姉瑀','歐陽文智','蔡明憬']
SA_ALL     = ['林定緯','林適緯','劉珈微','陳建道','陳星佑','張姉瑀','歐陽文智','蔡明憬']
C1_ALL     = ['林定緯','林適緯','劉珈微','陳建道']
C2_ALL     = ['陳星佑','蔡明憬','張姉瑀','歐陽文智']
COMPANY    = '巧克力汽車商業股份有限公司'

KPI_FILE_ID   = '1-8JmC6MlF-WPMsgeWqh43f_q5BiA3Q_G'
RENEW_FILE_ID = '1-6Wmly1lKSLVOEUspwcV9TPsghdjyMC_'
DAILY_FOLDER  = '1SDP7OJ79g6WqaoDqAQEeHZwj09MHTWyf'

SA_MAP = {'定緯':'林定緯','適緯':'林適緯','珈微':'劉珈微','建道':'陳建道',
          '星佑':'陳星佑','姉瑀':'張姉瑀','文智':'歐陽文智','明憬':'蔡明憬'}

MONTH_NAMES = ['一月','二月','三月','四月','五月','六月',
               '七月','八月','九月','十月','十一月','十二月']

def _safe_pct(n, d):
    return round(n/d, 4) if d and d > 0 else None

def _download(service, file_id):
    """下載為 bytes，用完即丟"""
    req = service.files().get_media(fileId=file_id)
    return req.execute()

def _latest_file(service, folder_id=None, name_contains=None):
    """Drive v3 搜尋最新檔案"""
    parts = ["trashed=false"]
    if folder_id:
        parts.append(f"'{folder_id}' in parents")
    if name_contains:
        parts.append(f"name contains '{name_contains}'")
    parts.append("mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'")
    result = service.files().list(
        q=" and ".join(parts),
        orderBy='modifiedTime desc',
        pageSize=1,
        fields='files(id,name)'
    ).execute()
    items = result.get('files', [])
    return (items[0]['id'], items[0]['name']) if items else (None, None)

# ─────────────────────────────────────────
# XML 解析核心
# ─────────────────────────────────────────
def _parse_xlsx_sheets(xlsx_bytes, target_sheets):
    """
    用 zipfile 直接讀 xlsx，只解析指定 sheet
    回傳：{ sheet_name: [ [cell_val, ...], ... ] }
    """
    import xml.etree.ElementTree as ET

    result = {}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
        names = zf.namelist()

        # 讀 sharedStrings
        shared = []
        if 'xl/sharedStrings.xml' in names:
            with zf.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in tree.findall('.//ns:si', ns):
                    t_nodes = si.findall('.//ns:t', ns)
                    shared.append(''.join(t.text or '' for t in t_nodes))

        # 讀 workbook.xml 取得 sheet name → rId 對應
        sheet_map = {}  # sheet_name → sheet xml path
        if 'xl/workbook.xml' in names:
            with zf.open('xl/workbook.xml') as f:
                tree = ET.parse(f)
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
                for sh in tree.findall('.//ns:sheet', ns):
                    sname = sh.get('name','')
                    rid   = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id','')
                    sheet_map[sname] = rid

        # 讀 workbook.xml.rels 取得 rId → 路徑
        rid_to_path = {}
        rels_path = 'xl/_rels/workbook.xml.rels'
        if rels_path in names:
            with zf.open(rels_path) as f:
                tree = ET.parse(f)
                for rel in tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rid_to_path[rel.get('Id','')] = 'xl/' + rel.get('Target','').lstrip('/')

        def parse_cell(c, shared):
            t = c.get('t','')
            v_node = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if v_node is None: return None
            val = v_node.text or ''
            if t == 's':
                try: return shared[int(val)]
                except: return val
            try: return float(val) if '.' in val else int(val)
            except: return val

        # 解析目標 sheets
        for sname in target_sheets:
            if sname not in sheet_map: continue
            rid = sheet_map[sname]
            path = rid_to_path.get(rid)
            if not path or path not in names: continue

            rows_data = []
            with zf.open(path) as f:
                tree = ET.parse(f)
                ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                for row_el in tree.findall(f'.//{{{ns}}}row'):
                    row = []
                    cells = row_el.findall(f'{{{ns}}}c')
                    if not cells:
                        rows_data.append([])
                        continue
                    # 計算最大 col index
                    import re
                    def col_idx(ref):
                        m = re.match(r'([A-Z]+)', ref or '')
                        if not m: return 0
                        s = m.group(1)
                        r = 0
                        for ch in s: r = r*26 + (ord(ch)-64)
                        return r-1

                    max_col = col_idx(cells[-1].get('r',''))
                    row = [None] * (max_col+1)
                    for c in cells:
                        ci = col_idx(c.get('r',''))
                        row[ci] = parse_cell(c, shared)
                    rows_data.append(row)
            result[sname] = rows_data

    return result


# ─────────────────────────────────────────
# 1. 日報表
# ─────────────────────────────────────────
def read_daily(service):
    now = datetime.now(TZ)
    curr_m = now.month

    file_id, title = _latest_file(service, DAILY_FOLDER, '歸仁日報表')
    if not file_id: return {}

    raw = _download(service, file_id)
    # 只需要 115年度 sheet
    sheets = _parse_xlsx_sheets(raw, ['115年度'])
    del raw

    ws = sheets.get('115年度', [])
    ytd={};  curr={}
    DAILY_MAP = {
        '林定緯':'林定緯','林適緯':'林適緯','劉珈微':'劉珈微','陳建道':'陳建道',
        '陳星佑':'陳星佑','蔡明憬':'蔡明憬','張姉瑀':'張姉瑀','歐陽文智':'歐陽文智',
        '一課合計':'歸仁一課','二課合計':'歸仁二課','總月累':'歸仁據點'
    }
    for row in ws:
        if not row: continue
        name = str(row[0] or '').strip()
        if name not in DAILY_MAP: continue
        key = DAILY_MAP[name]
        def g(i): return int(row[i]) if i<len(row) and isinstance(row[i],(int,float)) else 0
        # 年度累計: col 25=ytd_ord, 26=ytd_reg
        ytd[key]  = {'ord': g(25), 'reg': g(26)}
        # 當月: col = 1+(m-1)*2, 2+(m-1)*2
        cm_ord_c = 1+(curr_m-1)*2;  cm_reg_c = 2+(curr_m-1)*2
        curr[key] = {'ord': g(cm_ord_c), 'reg': g(cm_reg_c)}

    return {'ytd':ytd,'curr':curr,'curr_month':curr_m,'last_updated':title}


# ─────────────────────────────────────────
# 2. 週邊指標
# ─────────────────────────────────────────
def read_kpi(service, curr_month):
    raw = _download(service, KPI_FILE_ID)
    target = ['2026.Q1','2026.Q2', f'2026.{curr_month:02d}月']
    sheets = _parse_xlsx_sheets(raw, target)
    del raw

    sa_data = defaultdict(lambda: defaultdict(lambda:{
        'reg':0,'acc':0,'yi':0,'bing':0,'full':0,'loan':0,
        'rental':0,'no_ins':0,'base':0,'prem':0
    }))

    period_map = {'2026.Q1':'Q1','2026.Q2':'Q2'}

    for sname, rows in sheets.items():
        period = period_map.get(sname)
        if not period:
            try:
                m = int(sname.replace('2026.','').replace('月',''))
                period = f'M{m}'
            except: continue

        for row in rows:
            if not row or len(row)<7: continue
            sa_abbr  = str(row[1] or '').strip()
            customer = str(row[2] or '').strip()
            acc      = row[3]
            loan_co  = str(row[4] or '').strip()
            ins_type = str(row[6] or '').strip()
            premium  = row[7] if len(row)>7 else None

            if sa_abbr not in SA_MAP: continue
            sa = SA_MAP[sa_abbr]
            if COMPANY in customer: continue

            is_rental = (customer=='租車')
            sa_data[sa][period]['reg'] += 1
            if isinstance(acc,(int,float)): sa_data[sa][period]['acc'] += acc
            if isinstance(premium,(int,float)): sa_data[sa][period]['prem'] += premium
            if is_rental: sa_data[sa][period]['rental']+=1; continue

            sa_data[sa][period]['base'] += 1
            if loan_co=='元大': sa_data[sa][period]['loan'] += 1
            if ins_type=='乙式': sa_data[sa][period]['yi']+=1; sa_data[sa][period]['full']+=1
            elif ins_type=='丙式': sa_data[sa][period]['bing']+=1; sa_data[sa][period]['full']+=1
            elif ins_type=='甲式': sa_data[sa][period]['full']+=1
            else: sa_data[sa][period]['no_ins']+=1

    result = {}
    for sa in SA_ALL:
        # YTD = Q1+Q2（若7月以後再加M{curr}）
        ytd = {k:0 for k in ['reg','base','full','yi','loan','acc','prem']}
        for qp in ['Q1','Q2']:
            d = sa_data[sa][qp]
            for k in ytd: ytd[k] += d[k]
        if curr_month > 6:
            d = sa_data[sa][f'M{curr_month}']
            for k in ytd: ytd[k] += d[k]

        curr = dict(sa_data[sa][f'M{curr_month}'])

        def mk(d):
            reg=d.get('reg',0); base=d.get('base',0)
            acc=d.get('acc',0); full=d.get('full',0); yi=d.get('yi',0); loan=d.get('loan',0)
            return {'reg':reg,'base':base,'acc_t':int(acc),'full':full,'yi':yi,'loan':loan,
                    'acc_per':round(acc/reg) if reg>0 else 0,
                    '全險比':_safe_pct(full,base),'乙式比':_safe_pct(yi,base),
                    '分期比':_safe_pct(loan,base)}
        result[sa] = {'ytd':mk(ytd),'curr':mk(curr)}

    return result


# ─────────────────────────────────────────
# 3. 續保進度表
# ─────────────────────────────────────────
def read_renew(service, curr_month):
    raw = _download(service, RENEW_FILE_ID)
    month_sheet = f'115.{curr_month:02d}月續保'
    target = [month_sheet,'2026續保成績','2026首年車體續保','2026首年續保']
    sheets = _parse_xlsx_sheets(raw, target)
    del raw

    curr_data={}
    ws_curr = sheets.get(month_sheet,[])
    SA_RENEW = set(SA_ALL+['歸仁一課','歸仁二課','合計','劉宗鑫'])
    for row in ws_curr:
        if not row: continue
        sa=str(row[0] or '').strip()
        if sa not in SA_RENEW: continue
        def g(i): return row[i] if i<len(row) and isinstance(row[i],(int,float)) else 0
        curr_data[sa]={
            '整體':    {'den':g(1),'num':g(3),'rate':_safe_pct(g(3),g(1))},
            '首年':    {'den':g(7),'num':g(9),'rate':_safe_pct(g(9),g(7))},
            '首年車體':{'den':g(13),'num':g(15),'rate':_safe_pct(g(15),g(13))},
        }

    ytd_data={}
    completed = list(range(1, curr_month))
    for section,sname in [('整體','2026續保成績'),('首年車體','2026首年車體續保'),('首年','2026首年續保')]:
        ws=sheets.get(sname,[])
        for row in ws:
            if not row: continue
            sa=str(row[0] or '').strip()
            if sa not in SA_RENEW: continue
            n=d=0
            for m in completed:
                dc=8+(m-1)*2; nc=9+(m-1)*2
                if dc<len(row): d+=int(row[dc] or 0)
                if nc<len(row): n+=int(row[nc] or 0)
            if sa not in ytd_data: ytd_data[sa]={}
            ytd_data[sa][section]={'den':d,'num':n,'rate':_safe_pct(n,d)}

    return {'curr':curr_data,'ytd':ytd_data}


# ─────────────────────────────────────────
# 4. 有望客
# ─────────────────────────────────────────
def read_prospect(service):
    file_id,_ = _latest_file(service, name_contains='歸仁有望客名單')
    if not file_id: return {}

    raw = _download(service, file_id)
    sheets = _parse_xlsx_sheets(raw, ['750000'])
    del raw

    SUCCESS={'訂單','領牌','交車','退購'}
    EXCL={'來廠','外展'}
    result=defaultdict(lambda:{'walk':0,'walkC':0,'inv':0,'invC':0})

    ws=sheets.get('750000',[])
    for i,row in enumerate(ws):
        if i==0 or not row or len(row)<25: continue
        sa=row[5]; status=row[6]; source=row[14]; media=row[15]
        if not sa or sa not in SA_ALL: continue
        is_ok=status in SUCCESS
        if source=='自然來店':
            result[sa]['walk']+=1
            if is_ok: result[sa]['walkC']+=1
        elif source=='邀約來店' and str(media or '') not in EXCL:
            result[sa]['inv']+=1
            if is_ok: result[sa]['invC']+=1
    return dict(result)


# ─────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────
def get_guiren_kpi(service):
    now=datetime.now(TZ)
    curr_month=now.month

    daily   = read_daily(service)
    kpi     = read_kpi(service, curr_month)
    renew   = read_renew(service, curr_month)
    prospect= read_prospect(service)

    ytd_map  = daily.get('ytd',{})
    curr_map = daily.get('curr',{})
    hsin_pros={'Q2':{'walk':0,'walkC':0,'inv':1,'invC':0}}

    def _sum_kpi(members, pkey):
        r=b=f=yi=ln=at=0
        for sa in members:
            d=kpi.get(sa,{}).get(pkey,{})
            r+=d.get('reg',0); b+=d.get('base',0); f+=d.get('full',0)
            yi+=d.get('yi',0); ln+=d.get('loan',0); at+=d.get('acc_t',0)
        return {'reg':r,'base':b,'full':f,'yi':yi,'loan':ln,'acc_t':at,
                'acc_per':round(at/r) if r>0 else 0,
                '全險比':_safe_pct(f,b),'乙式比':_safe_pct(yi,b),'分期比':_safe_pct(ln,b)}

    def _sum_pros(members, inc_hsin=False):
        wk=wkc=iv=ivc=0
        for sa in members:
            p=prospect.get(sa,{})
            wk+=p.get('walk',0); wkc+=p.get('walkC',0)
            iv+=p.get('inv',0);  ivc+=p.get('invC',0)
        if inc_hsin: iv+=1
        return {'walk':wk,'walkC':wkc,'inv':iv,'invC':ivc,
                'walk_rate':_safe_pct(wkc,wk),'inv_rate':_safe_pct(ivc,iv)}

    def _sum_renew(members, section, period):
        n=d=0
        src=renew.get(period,{})
        for sa in members:
            r=src.get(sa,{}).get(section,{})
            n+=r.get('num',0); d+=r.get('den',0)
        return {'num':n,'den':d,'rate':_safe_pct(n,d)}

    def build(members, key, is_depot=False):
        ytd_reg  = sum(ytd_map.get(sa,{}).get('reg',0) for sa in members)
        curr_reg = sum(curr_map.get(sa,{}).get('reg',0) for sa in members)
        ytd_ord  = sum(ytd_map.get(sa,{}).get('ord',0) for sa in members)
        curr_ord = sum(curr_map.get(sa,{}).get('ord',0) for sa in members)
        kq=_sum_kpi(members,'ytd'); kc=_sum_kpi(members,'curr')
        pros=_sum_pros(members, inc_hsin=is_depot)
        # 據點 renew key = '合計'
        rkey = '合計' if is_depot else key
        return {
            'ytd':{**kq,'reg':ytd_reg,'ord':ytd_ord,**pros},
            'curr':{**kc,'reg':curr_reg,'ord':curr_ord},
            'renew':{
                'curr':renew.get('curr',{}).get(rkey,{}),
                'ytd': renew.get('ytd',{}).get(rkey,{})
            }
        }

    entities={}
    entities['歸仁據點']=build(C1_ALL+C2_ALL,'歸仁據點',is_depot=True)
    entities['歸仁一課']=build(C1_ALL,'歸仁一課')
    entities['歸仁二課']=build(C2_ALL,'歸仁二課')
    for sa in SA_DISPLAY:
        entities[sa]=build([sa],sa)

    return {
        'meta':{'curr_month':curr_month,'curr_month_name':MONTH_NAMES[curr_month-1],
                'updated_at':now.strftime('%Y-%m-%d %H:%M'),
                'daily_source':daily.get('last_updated','')},
        'entities':entities
    }
