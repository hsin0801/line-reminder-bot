"""
11 月銷售達人排行榜（內促備註 4）
------------------------------------------------
規則：11 月份據點銷售達人賞（門檻 5 台，需含一台 CR-V e:HEV P），
      各課第一 5,000 元紅包（同台數比較 CR-V e:HEV 台數），據點第一再 +5,000，法人第一再 +10,000。
      → 旗子顯示累計：課第一 5,000／據點第一 10,000／法人第一 20,000
目前（月初）先只排名次、不套門檻：ENFORCE_THRESHOLD = False，之後要啟用改成 True。

台數來源：兩據點每日日報表「11月」分頁的個人訂單月累（儀表板快取 monthly_by_person）。
CR-V e:HEV / Prestige：crv_trip 的 11 月訂單逐日計數（by_month / by_month_p）。
LINE 小幫手（只有 owner 能用）：
  岱宏訂單+1 / 岱宏訂單P+1 → 即時回報（暫時），下一份日報表進來後自動失效、以日報表為準
  以下為永久修正（用於無法當月領牌的訂單）：
  岱宏現訂交-1   → 總台數 -1
  岱宏現訂交P-1  → 總台數、CR-V e:HEV、Prestige 各 -1
  （+1 也可以）  岱宏取消修正 → 刪掉岱宏最近一筆現訂交修正（撤銷）
  達人排行 → 查目前名次
紀錄存在 Drive：sales_star_manual.json（日報表資料夾）
"""
import re
import threading
from datetime import datetime, timezone, timedelta

TPE = timezone(timedelta(hours=8))
MONTH = "11月"
MONTH_NO = "11"
THRESHOLD = 5
ENFORCE_THRESHOLD = False
PRIZE = {"dept": 5000, "branch": 5000, "faren": 10000}
DEPT_SHORT = {'永康一課': '永一', '永康二課': '永二', '永康三課': '永三', '歸仁一課': '歸一', '歸仁二課': '歸二'}
MANUAL_FILENAME = "sales_star_manual.json"

_cache = {"entries": None}
_lock = threading.Lock()


def _alias():
    from crv_trip import ALIAS
    return ALIAS


def _pat():
    names = '|'.join(sorted(_alias(), key=len, reverse=True))
    return re.compile(rf"^({names})(現訂交|訂單)(P)?([+＋\-－])(\d{{0,2}})台?$|^(達人排行|銷售達人)$|^({names})(取消修正|撤銷修正|取消現訂交)$", re.I)


def _norm(text):
    return text.replace(" ", "").replace("　", "").upper().replace("Ｐ", "P")


def _lines(text):
    return [l for l in (_norm(x) for x in text.splitlines()) if l]


def match(text):
    # 一則訊息可以多行，一行一筆（例：霽杰訂單+1↵適緯訂單P+1），每一行都要是指令才接手
    ls = _lines(text)
    pat = _pat()
    return bool(ls) and all(pat.match(l) for l in ls)


def _now():
    return datetime.now(TPE).replace(tzinfo=None).isoformat(timespec="seconds")


def load(force=False):
    if _cache["entries"] is None or force:
        from drive_json_store import load_json_from_drive
        from drive_reader import DAILY_REPORT_FOLDER_ID
        data = load_json_from_drive(DAILY_REPORT_FOLDER_ID, MANUAL_FILENAME) or {}
        _cache["entries"] = list(data.get("entries", []))
    return list(_cache["entries"])


def _add(name, delta, p, temp=False):
    from drive_json_store import save_json_to_drive
    from drive_reader import DAILY_REPORT_FOLDER_ID
    with _lock:
        entries = load()
        e = {"ts": _now(), "name": name, "delta": delta, "p": bool(p)}
        if temp:
            e["temp"] = True
        entries.append(e)
        save_json_to_drive(DAILY_REPORT_FOLDER_ID, MANUAL_FILENAME, {"entries": entries})
        _cache["entries"] = entries


def _undo(name):
    """刪掉此人最近一筆永久修正（現訂交±n），回傳被刪的那筆；沒有就回 None。"""
    from drive_json_store import save_json_to_drive
    from drive_reader import DAILY_REPORT_FOLDER_ID
    with _lock:
        entries = load(force=True)
        for i in range(len(entries) - 1, -1, -1):
            e = entries[i]
            if e.get("name") == name and not e.get("temp"):
                del entries[i]
                save_json_to_drive(DAILY_REPORT_FOLDER_ID, MANUAL_FILENAME, {"entries": entries})
                _cache["entries"] = entries
                return e
    return None


def _branch_cache(branch):
    import os, json
    if branch == "gueiren":
        from dashboard_parser import DATA_FILE
    else:
        from yongkang_parser import DATA_FILE
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, encoding="utf-8") as f:
            return json.load(f) or {}
    from drive_json_store import restore_cache
    return restore_cache(DATA_FILE) or {}


def standings():
    """[{name, branch, dept, short, total, ehev, p, adj, flags:[...], prize}]，已排序（法人名次）。"""
    from crv_trip import BRANCHES
    rows = []
    for b, cfg in BRANCHES.items():
        d = _branch_cache(b)
        months = d.get("monthly_months") or []
        idx = months.index(MONTH) if MONTH in months else None
        mbp = d.get("monthly_by_person") or {}
        crv = d.get("crv_trip") or {}
        rep = (crv.get("report_date") or "", (d.get("updated_at") or "")[:19])
        for nm, dept in cfg["dept"].items():
            tot = (mbp.get(nm, {}).get("訂單") or [])
            tot = tot[idx] if idx is not None and idx < len(tot) else 0
            rows.append({"name": nm, "branch": b, "dept": dept, "short": DEPT_SHORT.get(dept, dept),
                         "total": int(tot or 0),
                         "ehev": int(((crv.get("by_month") or {}).get(nm) or {}).get(MONTH_NO, 0)),
                         "p": int(((crv.get("by_month_p") or {}).get(nm) or {}).get(MONTH_NO, 0)),
                         "adj": 0, "pending": 0, "_rep": rep})
    by = {r["name"]: r for r in rows}
    for e in load():
        r = by.get(e.get("name"))
        if not r:
            continue
        if e.get("temp"):
            # 「訂單±n」即時回報：日報表（資料日期＋解析時間）比這筆新就不再計入，以日報表為準
            rd, upd = r["_rep"]
            day = e["ts"][:10]
            if not (day > rd or (day == rd and e["ts"] > upd)):
                continue
            r["pending"] += e["delta"]
        else:
            r["adj"] += e["delta"]
        r["total"] += e["delta"]
        if e.get("p"):
            r["ehev"] += e["delta"]
            r["p"] += e["delta"]
    for r in rows:
        r.pop("_rep", None)
        r["total"], r["ehev"], r["p"] = max(0, r["total"]), max(0, r["ehev"]), max(0, r["p"])
        r["ok"] = r["total"] >= THRESHOLD and r["p"] >= 1
    rows.sort(key=lambda r: (-r["total"], -r["ehev"]))

    def eligible(r):
        return r["total"] > 0 and (r["ok"] or not ENFORCE_THRESHOLD)

    for r in rows:
        r["flags"], r["prize"] = [], 0

    def award(group, label, amount_key):
        for r in rows:
            if group(r) and eligible(r):
                r["flags"].append(label)
                r["prize"] += PRIZE[amount_key]
                return

    for dept, short in DEPT_SHORT.items():
        award(lambda r, d=dept: r["dept"] == d, f"{short}第一", "dept")
    award(lambda r: r["branch"] == "yongkang", "永康第一", "branch")
    award(lambda r: r["branch"] == "gueiren", "歸仁第一", "branch")
    award(lambda r: True, "法人第一", "faren")
    return rows


def handle(text):
    ls = _lines(text)
    if len(ls) <= 1:
        return _handle_one(ls[0] if ls else text)
    out = []
    for l in ls:
        out.append(_handle_one(l, brief=True))
    try:
        rows = standings()
        top = "、".join(f"{r['name']} {r['total']}" for r in rows[:3])
        out.append(f"🏁 目前前三：{top}")
    except Exception:
        pass
    return "\n".join(out)


def _handle_one(text, brief=False):
    m = _pat().match(_norm(text))
    lines = []
    if m and m.group(7):
        name = _alias()[m.group(7)]
        try:
            e = _undo(name)
        except Exception as ex:
            return f"取消失敗，請稍後再試（{str(ex)[:60]}）"
        if not e:
            return f"{name} 沒有可以取消的現訂交修正紀錄。"
        lines.append(f"↩️ 已取消 {name} 的現訂交修正 {e['delta']:+d}" + ("（含 P）" if e.get("p") else "")
                     + f"（{e['ts'][5:16].replace('T', ' ')} 記錄的那筆）")
        m = None
        m_name = name
    else:
        m_name = None
    if m and m.group(1):
        name = _alias()[m.group(1)]
        temp, p, sign, num = m.group(2) == "訂單", bool(m.group(3)), m.group(4), m.group(5)
        n = int(num) if num else 1
        if n == 0 or n > 5:
            return "數字怪怪的，一次請調整 1～5 台，例如：岱宏現訂交-1"
        delta = n if sign in "+＋" else -n
        try:
            _add(name, delta, p, temp)
        except Exception as e:
            return f"記錄失敗，請稍後再試（{str(e)[:60]}）"
        lines.append(f"✅ 已記錄 {name} {'訂單' if temp else '現訂交'} {delta:+d}" + ("（含 CR-V e:HEV P）" if p else "")
                     + ("｜即時回報，下一份日報表進來後以日報表為準" if temp else "｜永久修正"))
    if brief and lines:
        return "\n".join(lines)
    try:
        rows = standings()
    except Exception as e:
        print(f"[sales_star] standings error: {e}")
        lines.append("目前還讀不到儀表板資料，紀錄已存好，晚點會反映。")
        return "\n".join(lines)
    if (m and m.group(1)) or m_name:
        name = m_name or _alias()[m.group(1)]
        i, r = next((i, r) for i, r in enumerate(rows, 1) if r["name"] == name)
        lines.append(f"🏁 {name}（{r['short']}）11 月 {r['total']} 台｜e:HEV {r['ehev']}｜P {r['p']}｜法人第 {i} 名")
        if r["flags"]:
            lines.append("🚩 目前拿旗：" + "、".join(r["flags"]) + f"（{r['prize']:,} 元）")
        return "\n".join(lines)
    lines.append(f"🏁 11 月銷售達人排行（門檻 {THRESHOLD} 台＋e:HEV P" + ("" if ENFORCE_THRESHOLD else "，目前先不套門檻") + "）")
    for i, r in enumerate(rows[:10], 1):
        flag = f" 🚩{r['flags'][-1]} {r['prize']:,}" if r["flags"] else ""
        lines.append(f"{i}. {r['name']}（{r['short']}）{r['total']} 台｜P {r['p']}{flag}")
    return "\n".join(lines)


DAILY_CRV_KEYS = ('CR-V(PET)', 'CR-V(e:HEV)')   # 歷史檔裡這兩個是「當日」數字，其餘車型是月累


_daily_cache = {"t": 0, "v": None}


def daily_series():
    """開場重播用（10 分鐘快取）：11 月每份日報表當天的個人訂單月累 {dates, series}。"""
    import time
    if _daily_cache["v"] is not None and time.time() - _daily_cache["t"] < 600:
        return _daily_cache["v"]
    v = _daily_series()
    _daily_cache.update(t=time.time(), v=v)
    return v


def _daily_series():
    import dashboard_parser, yongkang_parser
    from crv_trip import BRANCHES
    prefix = f"2026-{int(MONTH_NO):02d}-"
    per_day = {}
    for mod in (dashboard_parser, yongkang_parser):
        try:
            hist = mod.load_order_history()
        except Exception as e:
            print(f"[sales_star] history {mod.__name__}: {e}")
            continue
        run = {}
        for d in sorted(k for k in hist if k.startswith(prefix)):
            day = per_day.setdefault(d, {})
            for nm, models in hist[d].items():
                run[nm] = run.get(nm, 0) + sum(models.get(k, 0) or 0 for k in DAILY_CRV_KEYS)
                day[nm] = sum(v or 0 for k, v in models.items() if k not in DAILY_CRV_KEYS) + run[nm]
    people = [p for c in BRANCHES.values() for p in c["dept"]]
    dates = sorted(per_day)
    series = {}
    for p in people:
        v = [per_day[d].get(p, 0) for d in dates]
        for i in range(len(v) - 2, -1, -1):   # 月累只會增加：歷史檔若有異常偏高的日子，壓回後一天的數字
            v[i] = min(v[i], v[i + 1])
        series[p] = v
    return {"dates": dates, "series": series}
