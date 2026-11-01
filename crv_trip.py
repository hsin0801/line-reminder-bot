"""
觸發任務「CR-V e:HEV 保送出國」（歸仁＋永康）
------------------------------------------------
規則（宗鑫 2026-11-01 定）：
  9/1–12/31 CR-V e:HEV（S＋Prestige）累計 12 台即保送出國，每人各自計算。
  9–10 月：算「領牌」（已結算，寫死在 BASE，來源＝9/1–10/30 每日日報表逐日加總，與各課月累核對一致）
  11–12 月：算「訂單」，每份日報表進來自動累加
  另可用 LINE 小幫手補：
    定緯 CRV+1 / 定緯 CRV-1      → 即時回報，該據點下一份日報表進來後自動作廢（以日報表為準，不重複）
    定緯 CRV修正+1 / 修正-1       → 永久修正（日報表漏記、退訂、車型登錯等）
    CRV進度                       → 查兩據點全員進度
  只有宗鑫（ALARM_OWNER_ID）能用。

日報表「N月」分頁：每個人的 CR-V 車型欄（e:HEV S / Prestige）是「當日」數字，CR-V「月累」是當月累計。
常有前一天沒清掉、或月累晚一天才跟上的情形，所以：
  當天計數 = min(當日 e:HEV, CR-V 月累增加量)
  若當日有 e:HEV 但月累沒跟上 → 先記為「待確認」，之後幾天內月累增加、且當天沒有任何 CR-V 車型登記時補上（最多等 3 份）
（用 9/1–10/30 兩據點日報表驗證：領牌與各課月累 100% 一致）
進度存在 Drive：crv_trip_history.json（歸仁）／crv_trip_yk_history.json（永康）；LINE 紀錄：crv_trip_manual.json
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

BRANCHES = {
    "gueiren": {
        "label": "歸仁",
        "file_kw": "歸仁日報表115",
        "history": "crv_trip_history.json",
        "dept": {'林定緯': '歸仁一課', '林適緯': '歸仁一課', '陳建道': '歸仁一課',
                 '陳星佑': '歸仁二課', '蔡明憬': '歸仁二課', '張姉瑀': '歸仁二課', '歐陽文智': '歸仁二課'},
        # 9、10 月 CR-V e:HEV 領牌（9 月 一課7/二課11、10 月 一課4/二課5，與課月累一致）
        "base": {
            '林定緯': (5, 2), '林適緯': (2, 0), '陳建道': (0, 2), '陳星佑': (3, 0),
            '蔡明憬': (3, 0), '張姉瑀': (3, 2), '歐陽文智': (2, 3),
        },
    },
    "yongkang": {
        "label": "永康",
        "file_kw": "永康日報表",
        "history": "crv_trip_yk_history.json",
        "dept": {'謝岱宏': '永康一課', '林昭吾': '永康一課', '林明昌': '永康一課', '陳靖玟': '永康一課', '林駿軒': '永康一課',
                 '洪嘉綺': '永康二課', '王少佟': '永康二課', '李霽杰': '永康二課', '陳冠甫': '永康二課', '莊華昇': '永康二課',
                 '蘇嵩閔': '永康三課', '陳建志': '永康三課', '黃湘婷': '永康三課', '江念澤': '永康三課', '王裕維': '永康三課'},
        # 9 月 一課7/二課7/三課6（含已離職林祐諄 1 台，不列個人）、10 月 一課6/二課3/三課7，與課月累一致
        "base": {
            '謝岱宏': (3, 2), '林昭吾': (3, 1), '林明昌': (0, 1), '陳靖玟': (1, 1), '林駿軒': (0, 1),
            '洪嘉綺': (2, 0), '王少佟': (2, 0), '李霽杰': (1, 2), '陳冠甫': (1, 1), '莊華昇': (0, 0),
            '蘇嵩閔': (2, 1), '陳建志': (1, 0), '黃湘婷': (1, 2), '江念澤': (1, 1), '王裕維': (1, 3),
        },
    },
}
MANUAL_FILENAME = "crv_trip_manual.json"

ALIAS = {}
NAME_BRANCH = {}
for _b, _c in BRANCHES.items():
    for _p in _c["dept"]:
        ALIAS[_p] = _p
        ALIAS[_p[1:]] = _p          # 兩字簡稱：定緯、岱宏…（歐陽文智 → 陽文智 也可）
        NAME_BRANCH[_p] = _b
ALIAS['文智'] = '歐陽文智'


def _now():
    return datetime.now(TPE).replace(tzinfo=None).isoformat(timespec="seconds")


# ───────────────────────── 日報表解析 ─────────────────────────
def _norm(v):
    return str(v or '').upper().replace(':', '').replace(' ', '').replace('\n', '').replace('-', '')


def crv_cols(rows, section):
    """section='reg'（領牌）或 'order'（訂單）→ (CR-V 群組第一欄, e:HEV S 欄, Prestige 欄, 月累欄)"""
    r0 = oc = None
    for r in range(min(6, len(rows))):
        for c, v in enumerate(rows[r]):
            if _norm(v) == '訂單數':
                r0, oc = r, c
                break
        if r0 is not None:
            break
    if r0 is None or r0 + 2 >= len(rows):
        return None
    mr, tr = rows[r0 + 1], rows[r0 + 2]
    lo, hi = (1, oc) if section == 'reg' else (oc, oc + 40)
    heads = [c for c in range(lo, min(hi, len(mr))) if _norm(mr[c])]
    for i, c in enumerate(heads):
        if _norm(mr[c]) != 'CRV':
            continue
        end = heads[i + 1] if i + 1 < len(heads) else min(hi, len(tr))
        s = p = cum = None
        for k in range(c, end):
            L = _norm(tr[k]) if k < len(tr) else ''
            if s is None and (L == 'ES' or 'HEVS' in L):
                s = k
            elif p is None and (L == 'EP' or 'HEVP' in L):
                p = k
            elif cum is None and '月累' in L:
                cum = k
        if None not in (s, p, cum):
            return c, s, p, cum
    return None


def parse_sheet(wb, month, section, people):
    """{姓名: (當日 e:HEV, 當日 CR-V 各車型合計, CR-V 個人月累)}"""
    names = wb.sheetnames if hasattr(wb, 'sheetnames') else []
    sn = next((s for s in names if str(s).strip() == f'{month}月'), f'{month}月')
    rows = [list(r) for r in wb[sn].iter_rows(values_only=True)]
    cols = crv_cols(rows, section)
    if not cols:
        raise ValueError(f'{month}月 找不到 CR-V e:HEV 欄位')
    first, s, p, cum = cols
    num = lambda row, c: row[c] if c < len(row) and isinstance(row[c], (int, float)) else 0
    out = {}
    for row in rows:
        if not row:
            continue
        nm = str(row[0] or '').strip()
        if nm in people and nm not in out:
            out[nm] = (num(row, s) + num(row, p), sum(num(row, k) for k in range(first, cum)), num(row, cum))
    return out


def step(state, per):
    """依一份日報表更新每人狀態，回傳當天計數 {姓名: 台數}。"""
    inc = {}
    for nm, (ehev, allcrv, cum) in per.items():
        st = state.setdefault(nm, {'cum': 0, 'owed': 0, 'age': 0})
        d = max(0, cum - st['cum'])
        c = min(ehev, d)
        surplus = max(0, d - allcrv)          # 月累增加、但當天沒有任何 CR-V 車型登記 → 補前幾天的待確認
        pay = min(st['owed'], surplus)
        if ehev > c:
            st['owed'] += ehev - c
            st['age'] = 0
        st['owed'] -= pay
        if st['owed'] > 0:
            st['age'] += 1
            if st['age'] > 3:
                st['owed'], st['age'] = 0, 0
        k = int(round(c + pay))
        if k > 0:
            inc[nm] = k
        st['cum'] = cum
    return inc


def update_history(branch, max_files=12):
    """把還沒處理過的 11/12 月日報表依日期順序累加進 history，回傳 history。"""
    from drive_reader import get_drive_service, download_file, DAILY_REPORT_FOLDER_ID
    from drive_json_store import load_json_from_drive, save_json_to_drive
    from dashboard_parser import _parse_filename_date, _open_wb

    cfg = BRANCHES[branch]
    people = list(cfg["dept"])
    hist = load_json_from_drive(DAILY_REPORT_FOLDER_ID, cfg["history"]) or {}
    for k in ("processed", "prev", "days"):
        hist.setdefault(k, {})
    year = int(CONFIG["end"][:4])
    service = get_drive_service()
    q = (f"'{DAILY_REPORT_FOLDER_ID}' in parents and name contains '{cfg['file_kw']}' "
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
            per = parse_sheet(wb, mm, 'order', people)
        finally:
            try:
                wb.close()
            except Exception:
                pass
        del content
        state = hist["prev"].setdefault(str(mm), {})
        if any(not isinstance(v, dict) for v in state.values()):
            state.clear()
        inc = step(state, per)
        dkey = f"{year}-{mm:02d}-{dd:02d}"
        if inc:
            day = hist["days"].setdefault(dkey, {})
            for nm, c in inc.items():
                day[nm] = day.get(nm, 0) + c
        hist["processed"][f["id"]] = {"name": f["name"], "date": dkey, "created": ct, "inc": inc}
        changed = True
    if changed:
        save_json_to_drive(DAILY_REPORT_FOLDER_ID, cfg["history"], hist)
    return hist


def build(branch="gueiren"):
    """給儀表板用的資料（存進各據點快取的 crv_trip 欄位）。"""
    cfg = BRANCHES[branch]
    people = list(cfg["dept"])
    hist = update_history(branch)
    by = {p: {'9': cfg["base"][p][0], '10': cfg["base"][p][1], '11': 0, '12': 0} for p in people}
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
    return {**CONFIG, "branch": branch, "label": cfg["label"], "people": people, "dept": cfg["dept"],
            "by_month": by,
            "report_date": last["date"] if last else "2026-10-30",
            "report_name": last["name"] if last else f"{cfg['label']}日報表 10/30",
            "report_created": rc or "2026-10-30T17:44:00"}


# ───────────────────────── LINE 小幫手 ─────────────────────────
_NAMES = '|'.join(sorted(ALIAS, key=len, reverse=True))
_PAT = re.compile(rf"^({_NAMES})CRV(修正)?([+＋\-－])(\d{{0,2}})台?$|^CRV進度$", re.I)
_cache = {"entries": None}
_lock = threading.Lock()


def _norm_text(text):
    return text.replace(" ", "").replace("　", "").upper().replace("CR-V", "CRV")


def match(text):
    return bool(_PAT.match(_norm_text(text)))


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


def _branch_data(branch):
    import os, json
    if branch == "gueiren":
        from dashboard_parser import DATA_FILE
    else:
        from yongkang_parser import DATA_FILE
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, encoding="utf-8") as f:
            return (json.load(f) or {}).get("crv_trip")
    from drive_json_store import restore_cache
    return (restore_cache(DATA_FILE) or {}).get("crv_trip")


def status(branch):
    """({姓名: {report, fix, pending, cur}}, data)，report＝日報表(9–10 領牌＋11–12 訂單)。"""
    data = _branch_data(branch)
    if not data or data.get("error") or not data.get("people"):
        return None, None
    rc = data.get("report_created") or ""
    out = {p: {"report": sum(data["by_month"][p].values()), "fix": 0, "pending": 0} for p in data["people"]}
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
    m = _PAT.match(_norm_text(text))
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
            st, data = status(NAME_BRANCH[name])
        except Exception as e:
            print(f"[crv_trip] status error: {e}")
            st = None
        if not st:
            lines.append("目前還讀不到儀表板資料，紀錄已存好，晚點會反映。")
            return "\n".join(lines)
        s = st[name]
        left = max(0, data["target"] - s["cur"])
        extra = (f"，修正 {s['fix']:+d}" if s["fix"] else "") + (f"，LINE 回報 {s['pending']:+d}" if s["pending"] else "")
        lines.append(f"✈️ {name} 目前 {s['cur']} 台（日報表 {s['report']} 台{extra}）")
        lines.append("🎉 已達標，保送出國！" if left == 0 else f"距保送出國還差 {left} 台 🔥")
        return "\n".join(lines)
    lines.append(f"✈️ CR-V e:HEV 保送出國（目標 {CONFIG['target']} 台）")
    for b in ("gueiren", "yongkang"):
        try:
            st, data = status(b)
        except Exception as e:
            print(f"[crv_trip] status error: {e}")
            st = None
        label = BRANCHES[b]["label"]
        if not st:
            lines.append(f"【{label}】讀不到資料")
            continue
        rd = data.get("report_date", "")
        lines.append(f"【{label}｜日報表 {int(rd[5:7])}/{int(rd[8:10])}】")
        for i, (p, s) in enumerate(sorted(st.items(), key=lambda kv: -kv[1]["cur"]), 1):
            left = max(0, data["target"] - s["cur"])
            tag = "達標 🎉" if left == 0 else f"差 {left}"
            lines.append(f"{i}. {p} {s['cur']} 台｜{tag}" + (" *含LINE" if s["pending"] or s["fix"] else ""))
    return "\n".join(lines)
