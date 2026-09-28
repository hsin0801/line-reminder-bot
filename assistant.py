"""
小幫手：用歸仁績效指標的真實資料回答問題
------------------------------------------------
資料只取 guiren_kpi_reader（戰情室儀表板同一份）。業績速報業務沒有權限看，不能接進來。
資料一天只更新一次，快取 30 分鐘，避免每次提問都去 Drive 下載（LINE 回覆有時限）。
"""

import os
import re
import time
import requests

CACHE_SECONDS = 1800
GROQ_MODEL = os.environ.get("GROQ_MODEL", "openai/gpt-oss-120b")

_cache = {"text": None, "ts": 0}


def _pct(v):
    return f"{v * 100:.1f}%" if isinstance(v, (int, float)) else "—"


def _rate(num, den):
    return f"{num / den * 100:.1f}%" if den else "—"


def _renew_line(label, r):
    if not r or not r.get("den"):
        return f"{label} —"
    num, den = int(r.get("num", 0)), int(r.get("den", 0))
    return f"{label} {num}/{den}（{_pct(r.get('rate'))}，未續 {den - num}）"


def _entity_block(name, e):
    c, y = e.get("curr", {}), e.get("ytd", {})
    rc, ry = e.get("renew", {}).get("curr", {}), e.get("renew", {}).get("ytd", {})
    lines = [
        f"【{name}】",
        f"本月：訂單 {c.get('ord', 0)}、領牌 {c.get('reg', 0)}（保險母數 {c.get('base', 0)}）、"
        f"全險比 {_pct(c.get('全險比'))}、乙式比 {_pct(c.get('乙式比'))}、分期比 {_pct(c.get('分期比'))}、"
        f"配件總額 {c.get('acc_t', 0):,} 元（每台 {c.get('acc_per', 0):,} 元）",
        f"今年累計：訂單 {y.get('ord', 0)}、領牌 {y.get('reg', 0)}、"
        f"全險 {y.get('full', 0)}/{y.get('base', 0)}（{_rate(y.get('full', 0), y.get('base', 0))}）、"
        f"乙式 {y.get('yi', 0)}/{y.get('base', 0)}（{_rate(y.get('yi', 0), y.get('base', 0))}）、"
        f"分期 {y.get('loan', 0)}/{y.get('base', 0)}（{_rate(y.get('loan', 0), y.get('base', 0))}）、"
        f"配件總額 {y.get('acc_t', 0):,} 元（每台 {y.get('acc_per', 0):,} 元）、"
        f"來店 {y.get('walk', 0)} 成交 {y.get('walkC', 0)}（{_rate(y.get('walkC', 0), y.get('walk', 0))}）、"
        f"邀約 {y.get('inv', 0)} 成交 {y.get('invC', 0)}（{_rate(y.get('invC', 0), y.get('inv', 0))}）",
        "本月續保：" + "、".join(_renew_line(k, rc.get(k)) for k in ("整體", "首年", "首年車體")),
        "今年續保：" + "、".join(_renew_line(k, ry.get(k)) for k in ("整體", "首年", "首年車體")),
    ]
    return "\n".join(lines)


def _rankings(entities, people):
    """模型排序常出錯，常見排名由程式先排好。"""
    def renew(n, period):
        return entities[n].get("renew", {}).get(period, {}).get("整體", {})

    def rank(title, key, fmt, reverse):
        rows = [(n, key(n)) for n in people]
        rows = [(n, v) for n, v in rows if v is not None]
        rows.sort(key=lambda x: x[1], reverse=reverse)
        return f"{title}：" + " > ".join(f"{n} {fmt(v)}" for n, v in rows)

    def renew_rate(p):
        return lambda n: renew(n, p).get("rate") if renew(n, p).get("den") else None

    def unrenewed(p):
        return lambda n: int(renew(n, p).get("den", 0) - renew(n, p).get("num", 0)) if renew(n, p).get("den") else None

    lines = [
        "【業務排名（程式已排好，回答排名問題請直接照這裡）】",
        rank("本月整體續保率 高→低", renew_rate("curr"), _pct, True),
        rank("本月未續台數 多→少", unrenewed("curr"), lambda v: f"{v}台", True),
        rank("今年整體續保率 高→低", renew_rate("ytd"), _pct, True),
        rank("本月領牌 多→少", lambda n: entities[n].get("curr", {}).get("reg"), lambda v: f"{v}台", True),
        rank("本月訂單 多→少", lambda n: entities[n].get("curr", {}).get("ord"), lambda v: f"{v}張", True),
        rank("今年領牌 多→少", lambda n: entities[n].get("ytd", {}).get("reg"), lambda v: f"{v}台", True),
        rank("今年每台配件 高→低", lambda n: entities[n].get("ytd", {}).get("acc_per"), lambda v: f"{v:,}元", True),
    ]
    return "\n".join(lines)


def format_kpi(data):
    meta = data.get("meta", {})
    entities = data.get("entities", {})
    groups = [n for n in ("歸仁據點", "歸仁一課", "歸仁二課") if n in entities]
    people = [n for n in entities if n not in groups]

    parts = [
        f"資料月份：{meta.get('curr_month_name', '')}（{meta.get('curr_month', '')}月），"
        f"資料來源 {meta.get('daily_source', '')}，更新於 {meta.get('updated_at', '')}",
        "說明：全險比=(乙式+丙式)÷保險母數；乙式比=乙式÷母數；分期比=元大分期÷母數；"
        "「未續」=續保母數減已續保台數。「歸仁據點」是全據點合計，一課、二課是課別合計。",
    ]
    parts += [_entity_block(n, entities[n]) for n in groups + people]
    parts.append(_rankings(entities, people))

    llc = data.get("llc", {})
    if llc.get("person"):
        rows = [f"{n} 成功 {v.get('成功', 0)}/保有客 {v.get('保有客', 0)}（{v.get('比率', 0)}%）"
                for n, v in llc["person"].items()]
        t = llc.get("total", {})
        rows.append(f"合計 成功 {t.get('成功', 0)}/保有客 {t.get('保有客', 0)}（{t.get('比率', 0)}%）")
        parts.append("【LLC】\n" + "\n".join(rows))
    return "\n\n".join(parts)


def get_context():
    now = time.time()
    if _cache["text"] and now - _cache["ts"] < CACHE_SECONDS:
        return _cache["text"]
    from guiren_kpi_reader import get_guiren_kpi
    from drive_reader import get_drive_service
    text = format_kpi(get_guiren_kpi(get_drive_service()))
    _cache["text"], _cache["ts"] = text, now
    return text


SYSTEM_PROMPT = """你是 Honda 歸仁營業所 LINE 群組的 AI 小幫手，用繁體中文回答，語氣輕鬆、可以帶點幽默，但要簡潔（通常 5 行以內）。
LINE 不支援 Markdown：不要用表格、粗體（**）、標題（#），列表用「・」或數字。

下面是歸仁營業所最新的績效資料。回答業績、保險、配件、來店、邀約、續保、LLC 相關問題時：
- 數字只能引用資料裡的，不可以自己估算或編造；需要比較或排名時，直接照資料裡的數字排
- 資料裡沒有的（例如業績速報、目標、客戶個資），就說查不到，不要猜
- 回答時提一下資料更新時間
和業績無關的閒聊就正常聊天。

=== 歸仁績效資料 ===
{context}"""

PROMO_PROMPT = """

=== 本月內促辦法 ===
{promo}

回答內促問題時注意：
- 績效資料只有「本月領牌數」，分不出是不是現訂交（當月訂、當月領牌），也分不出車款和直販。
  需要現訂交台數的條件（例如備註4、備註5），請用本月領牌數估算，並明講「以目前領牌數估算，不一定都是現訂交，月底以進度表為準」
- 需要車款別台數的條件（例如備註2、備註5），資料裡沒有車款，就說明無法判斷
- 配件（備註8）可以直接用資料裡的本月每台配件金額判斷"""

PROMO_KEYWORDS = ("內促", "備註", "獎金", "紅包", "獎勵", "現訂交", "達成賞", "扣款", "扣3000", "扣3,000")
PROMO_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "promo")


def get_promo():
    from datetime import datetime, timedelta, timezone
    month = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m")
    path = os.path.join(PROMO_DIR, f"{month}.txt")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return f.read()


# Groq 免費方案每個模型每分鐘 token 有上限，一題約 4k token；被限流(429)就換下一個模型
FALLBACK_MODELS = [GROQ_MODEL, "qwen/qwen3.8-27b"]
MODEL_PARAMS = {
    "openai/gpt-oss-120b": {"reasoning_effort": "low"},
    "qwen/qwen3.8-27b": {"reasoning_format": "hidden"},
}


def _call(model, messages):
    return requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {os.environ.get('GROQ_API_KEY')}",
            "Content-Type": "application/json",
        },
        json={"model": model, "temperature": 0.3, "messages": messages, **MODEL_PARAMS.get(model, {})},
        timeout=20,
    )


def answer(question, models=None):
    try:
        context = get_context()
    except Exception as e:
        print(f"[ASSISTANT] 讀取績效資料失敗: {e}")
        context = "（目前讀不到績效資料，被問到數字請說資料暫時讀不到）"

    system = SYSTEM_PROMPT.format(context=context)
    if any(k in question for k in PROMO_KEYWORDS):
        promo = get_promo()
        system += PROMO_PROMPT.format(promo=promo) if promo else "\n\n（本月內促辦法還沒建檔，被問到內促請說還查不到本月辦法）"
    return chat([
        {"role": "system", "content": system},
        {"role": "user", "content": question},
    ], models)


def chat(messages, models=None):
    for model in models or FALLBACK_MODELS:
        resp = _call(model, messages)
        if resp.status_code == 200:
            text = resp.json()["choices"][0]["message"]["content"]
            return re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()
        print(f"[ASSISTANT] {model} HTTP {resp.status_code}: {resp.text[:200]}")
        if resp.status_code != 429:
            break
    raise RuntimeError(f"Groq HTTP {resp.status_code}: {resp.text[:200]}")
