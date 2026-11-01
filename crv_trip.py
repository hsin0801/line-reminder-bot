"""
觸發任務「CR-V e:HEV 保送出國」
------------------------------------------------
規則（宗鑫 2026-11-01 定）：
  9/1–12/31 CR-V e:HEV（S＋Prestige）累計 12 台即保送出國，每人各自計算。
  9–10 月：算「領牌」（已結算，寫死在 BASE，來源＝9/1–10/30 每日日報表逐日加總，與各課月累核對一致）
  11–12 月：算「訂單」，每份日報表進來自動累加
  另可用 LINE 小幫手補：
    定緯 CRV+1 / 定緯 CRV-1      → 即時回報，下一份日報表進來後自動作廢（以日報表為準，不重複）
    定緯 CRV修正+1 / 修正-1       → 永久修正（日報表漏記、退訂、車型登錯等）
    CRV進度                       → 查全員進度
  只有宗鑫（ALARM_OWNER_ID）能用。

日報表「N月」分頁的訂單區：每個人的車型欄（e:HEVS / e:HEVP）是「當日」數字，
常有前一天沒清掉的情形，所以當天只認「CR-V 訂單個人月累」真的有增加的部分：
  當天計數 = min(當日 e:HEV 訂單, CR-V 月累增加量)
每份日報表只處理一次，進度存在 Drive：crv_trip_history.json；LINE 紀錄：crv_trip_manual.json
"""
import re
import threading
from datetime import datetime, timezone, timedelta

TPE = timezone(timedelta(hours=8))

CONFIG = {
    "title": "CR-V e:HEV 保送出國",
    "target": 12,
    "start": "2026-09-01",
    "end": "2026-12-31",
    "hide_from": "2027-01-01",
    "order_months": [11, 12],
}
PEOPLE = ['林定緯', '林適緯', '陳建道', '陳星佑', '蔡明憬', '張姉瑀', '歐陽文智']
DEPT = {'林定緯': '歸仁一課', '林適緯': '歸仁一課', '陳建道': '歸仁一課',
        '陳星佑': '歸仁二課', '蔡明憬': '歸仁二課', '張姉瑀': '歸仁二課', '歐陽文智': '歸仁二課'}
ALIAS = {'定緯': '林定緯', '適緯': '林適緯', '建道': '陳建道', '星佑': '陳星佑',
         '明憬': '蔡明憬', '姉瑀': '張姉瑀', '文智': '歐陽文智'}
for _p in PEOPLE:
    ALIAS[_p] = _p
# 9、10 月 CR-V e:HEV 領牌（日報表 9/1–10/30 逐日加總；9 月一課7/二課11、10 月一課4/二課5，與課月累一致）
BASE = {
    '林定緯': {'9': 5, '10': 2}, '林適緯': {'9': 2, '10': 0}, '陳建道': {'9': 0, '10': 2},
    '陳星佑': {'9': 3, '10': 0}, '蔡明憬': {'9': 3, '10': 0}, '張姉瑀': {'9': 3, '10': 2},
    '歐陽文智': {'9': 2, '10': 3},
}
HISTORY_FILENAME = "crv_trip_history.json"
MANUAL_FILENAME = "crv_trip_manual.json"


def _now():
    return datetime.now(TPE).replace(tzinfo=None).isoformat(timespec="seconds")


# ───────────────────────── 日報表 → 11/12 月 e:HEV 訂單 ─────────────────────────
def _order_cols(rows):
    """回傳 (hevs_col, hevp_col, cum_col)；找不到回傳 None。"""
    hr = oc = None
    for r in range(min(6, len(rows))):
        for c, v in enumerate(rows[r]):
            if v and '訂單數' in str(v):
                hr, oc = r, c
                break
        if hr is not None:
            break
    if hr is None:
        return None
    crv = None
    for r in range(hr, min(hr + 3, len(rows))):
        for c in range(oc, min(oc + 30, len(rows[r]))):
            if str(rows[r][c] or '').strip().upper().replace('-', '') == 'CRV':
                crv = (r, c)
                break
        if crv:
            break
    if not crv or crv[0] + 1 >= len(rows):
        return None
    sub = rows[crv[0] + 1]
    lab = lambda c: str(sub[c] or '').upper().replace(':', '').replace(' ', '').replace('\n', '') if c < len(sub) else ''
    s = p = cum = None
    for c in range(crv[1], crv[1] + 7):
        L = lab(c)
        if 'HEVS' in L and s is None:
            s = c
        elif 'HEVP' in L and p is None:
            p = c
        elif '月累' in L and cum is None:
            cum = c
    if s is None or p is None or cum is None:
        return None
    return s, p, cum


def parse_order_sheet(wb, month):
    """{姓名: (當日 e:HEV 訂單, CR-V 訂單個人月累)}"""
    ws = wb[f'{month}月']
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    cols = _order_cols(rows)
    if not cols:
        raise ValueError(f'{month}月 找不到 CR-V e:HEV 訂單欄')
    s, p, cum = cols
    num = lambda row, c: row[c] if c < len(row) and isinstance(row[c], (int, float)) else 0
    out = {}
    for row in rows:
        if not row:
            continue
        nm = str(row[0] or '').strip()
        if nm in PEOPLE and nm not in out:
            out[nm] = (num(row, s) + num(row, p), num(row, cum))
    return out


def update_history(max_files=12):
    """把還沒處理過的 11/12 月日報表依日期順序累加進 history，回傳 history。"""
    from drive_reader import get_drive_service, download_file, DAILY_REPORT_FOLDER_ID
    from drive_json_store import load_json_from_drive, save_json_to_drive
    from dashboard_parser import _parse_filename_date, _open_wb

    hist = load_json_from_drive(DAILY_REPORT_FOLDER_ID, HISTORY_FILENAME) or {}
    hist.setdefault("processed", {})
    hist.setdefault("prev", {})
    hist.setdefault("days", {})
    year = int(CONFIG["end"][:4])
    service = get_drive_service()
    q = (f"'{DAILY_REPORT_FOLDER_ID}' in parents and name contains '歸仁日報表115' "
         f"and trashed = false and createdTime > '{year}-10-25T00:00:00'")
    files = service.files().list(q=q, pageSize=200, fields="files(id, name, createdTime)").execute().get("files", [])
    todo = []
    for f in files:
        mm, dd = _parse_filename_date(f["name"])
        if mm in CONFIG["order_months"] and f["id"] not in hist["processed"]:
            todo.append(((mm, dd, f.get("createdTime", "")), f))
    todo.sort(key=lambda x: x[0])
    changed = False
    for (mm, dd, ct), f in todo[:max_files]:
        content = download_file(f["id"])
        wb = _open_wb(content, f["name"])
        try:
            per = parse_order_sheet(wb, mm)
        finally:
            try:
                wb.close()
            except Exception:
                pass
        del content
        prev = hist["prev"].setdefault(str(mm), {})
        dkey = f"{year}-{mm:02d}-{dd:02d}"
        inc = {}
        for nm, (daily, cum) in per.items():
            c = int(round(min(daily, max(0, cum - prev.get(nm, 0)))))
            if c > 0:
                inc[nm] = c
            prev[nm] = cum
        if inc:
            day = hist["days"].setdefault(dkey, {})
            for nm, c in inc.items():
                day[nm] = day.get(nm, 0) + c
        hist["processed"][f["id"]] = {"name": f["name"], "date": dkey, "created": ct, "inc": inc}
        changed = True
    if changed:
        save_json_to_drive(DAILY_REPORT_FOLDER_ID, HISTORY_FILENAME, hist)
    return hist


def build():
    """給儀表板用的資料（存進歸仁快取 crv_trip 欄位）。"""
    hist = update_history()
    by = {p: {'9': BASE[p]['9'], '10': BASE[p]['10'], '11': 0, '12': 0} for p in PEOPLE}
    for d, inc in hist.get("days", {}).items():
        m = str(int(d[5:7]))
        for nm, c in inc.items():
            if nm in by and m in by[nm]:
                by[nm][m] += c
    proc = list(hist.get("processed", {}).values())
    last = max(proc, key=lambda x: (x["date"], x.get("created", ""))) if proc else None
    rc = None
    if last and last.get("created"):
        rc = (datetime.fromisoformat(last["created"].replace("Z", "+00:00")).astimezone(TPE)
              .replace(tzinfo=None).isoformat(timespec="seconds"))
    return {**CONFIG, "people": PEOPLE, "dept": DEPT, "by_month": by,
            "report_date": last["date"] if last else "2026-10-30",
            "report_name": last["name"] if last else "歸仁日報表115 10 30",
            "report_created": rc or "2026-10-30T17:29:43"}


# ───────────────────────── LINE 小幫手 ─────────────────────────
_NAMES = '|'.join(sorted(ALIAS, key=len, reverse=True))
_PAT = re.compile(rf"^({_NAMES})CRV(修正)?([+＋\-－])(\d{{0,2}})台?$|^CRV進度$", re.I)
_cache = {"entries": None}
_lock = threading.Lock()


def _norm(text):
    return text.replace(" ", "").replace("　", "").upper().replace("CR-V", "CRV")


def match(text):
    return bool(_PAT.match(_norm(text)))


def load(force=False):
    if _cache["entries"] is None or force:
        from drive_json_store import load_json_from_drive
        from drive_reader import DAILY_REPORT_FOLDER_ID
        data = load_json_from_drive(DAILY_REPORT_FOLDER_ID, MANUAL_FILENAME) or {}
        _cache["entries"] = list(data.get("entries", []))
    return list(_cache["entries"])


def _add(name, delta, fix):
    from drive_json_store import save_json_to_drive
    from drive_reader import DAILY_REPORT_FOLDER_ID
    with _lock:
        entries = load()
        entries.append({"ts": _now(), "name": name, "delta": delta, "fix": bool(fix)})
        save_json_to_drive(DAILY_REPORT_FOLDER_ID, MANUAL_FILENAME, {"entries": entries})
        _cache["entries"] = entries


def status(data=None):
    """{姓名: {report, fix, pending, cur}}，report＝日報表(9–10 領牌＋11–12 訂單)。"""
    if data is None:
        import os, json
        from dashboard_parser import DATA_FILE
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, encoding="utf-8") as f:
                data = (json.load(f) or {}).get("crv_trip")
        else:
            from drive_json_store import restore_cache
            data = ((restore_cache(DATA_FILE) or {}).get("crv_trip"))
    if not data or data.get("error"):
        return None, None
    rc = data.get("report_created") or ""
    out = {}
    for p in data["people"]:
        rep = sum(data["by_month"][p].values())
        out[p] = {"report": rep, "fix": 0, "pending": 0}
    for e in load():
        if e.get("name") not in out:
            continue
        if e.get("fix"):
            out[e["name"]]["fix"] += e["delta"]
        elif e["ts"] > rc:
            out[e["name"]]["pending"] += e["delta"]
    for v in out.values():
        v["cur"] = v["report"] + v["fix"] + v["pending"]
    return out, data


def handle(text):
    m = _PAT.match(_norm(text))
    lines = []
    if m and m.group(1):
        name = ALIAS[m.group(1)]
        fix, sign, num = bool(m.group(2)), m.group(3), m.group(4)
        n = int(num) if num else 1
        if n == 0 or n > 5:
            return "數字怪怪的，一次請回報 1～5 台，例如：定緯 CRV+1"
        delta = n if sign in "+＋" else -n
        try:
            _add(name, delta, fix)
        except Exception as e:
            return f"記錄失敗，請稍後再試（{str(e)[:60]}）"
        kind = "永久修正" if fix else "即時回報（日報表進來後以日報表為準）"
        lines.append(f"✅ 已記錄 {name} CR-V e:HEV {delta:+d}｜{kind}")
    try:
        st, data = status()
    except Exception as e:
        print(f"[crv_trip] status error: {e}")
        st = None
    if not st:
        lines.append("目前還讀不到儀表板資料，紀錄已存好，晚點會反映。")
        return "\n".join(lines)
    tgt = data["target"]
    if m and m.group(1):
        name = ALIAS[m.group(1)]
        s = st[name]
        left = max(0, tgt - s["cur"])
        extra = (f"，修正 {s['fix']:+d}" if s["fix"] else "") + (f"，LINE 回報 {s['pending']:+d}" if s["pending"] else "")
        lines.append(f"✈️ {name} 目前 {s['cur']} 台（日報表 {s['report']} 台{extra}）")
        lines.append("🎉 已達標，保送出國！" if left == 0 else f"距保送出國還差 {left} 台 🔥")
    else:
        rd = data.get("report_date", "")
        lines.append(f"✈️ CR-V e:HEV 保送出國（目標 {tgt} 台｜日報表 {int(rd[5:7])}/{int(rd[8:10])}）")
        for i, (p, s) in enumerate(sorted(st.items(), key=lambda kv: -kv[1]["cur"]), 1):
            left = max(0, tgt - s["cur"])
            tag = "達標 🎉" if left == 0 else f"差 {left}"
            lines.append(f"{i}. {p} {s['cur']} 台｜{tag}" + (" *含LINE" if s["pending"] or s["fix"] else ""))
    return "\n".join(lines)
