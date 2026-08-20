#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
条件リスト銘柄のエントリー条件・撤退ラインを、依頼者フレームに従って機械的に再設計する。

設計順序は依頼者の制約どおり「構造 → 撤退ライン → サイズ」で固定してある。
サイズを先に決めてから撤退ラインを動かす経路はコード上存在しない。

  ① Q1(H1)進捗率の過去平均超過分
  ② 保守性倍率 = 前期実績 ÷ 前期期初予想
  ③ 信用需給（買残÷20日平均出来高 = 何日分 / 売残実株数 / 直近数週の増減）
  ④ 20日平均売買代金

使い方:
    python3 designer.py                 # 全銘柄
    python3 designer.py 8159 4368       # 銘柄指定
    python3 designer.py --json out.json # 生データもJSONで吐く

ネットワーク要件: query2.finance.yahoo.com / webapi.yanoshin.jp / finance.yahoo.co.jp /
tdnet-pdf.kabutan.jp へ到達できる環境で実行すること。
"""

import json
import math
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

try:
    import pdfplumber
except BaseException:
    # 環境によっては cryptography のネイティブ拡張が壊れており、
    # ImportError ではなく pyo3 の PanicException(BaseException 直系) で落ちる。
    # PDF 解析は補助機能なので、欠けても価格・信用の設計は続行させる。
    pdfplumber = None

JST = timezone(timedelta(hours=9))
SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = SCRIPT_DIR / "config.json"
CACHE_DIR = SCRIPT_DIR / ".cache"

IPHONE_UA = (
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1"
)
SLEEP = 0.4

UNSET = "未取得"


# ─────────────────────────── 取得層 ───────────────────────────

def _get(url, *, binary=False, timeout=30, tries=3):
    """iPhone UA 固定の GET。失敗は None を返す（例外で全体を落とさない）。"""
    headers = {"User-Agent": IPHONE_UA, "Accept": "*/*"}
    for attempt in range(tries):
        try:
            r = requests.get(url, headers=headers, timeout=timeout)
            if r.status_code == 200:
                return r.content if binary else r.text
            if r.status_code in (403, 404):
                return None
        except requests.RequestException:
            pass
        time.sleep(2 ** attempt)
    return None


def fetch_bars(code, rng="1y"):
    """日足 OHLCV。株式分割はYahoo側の調整済み系列を使い、分割イベントも併せて返す。"""
    url = (f"https://query2.finance.yahoo.com/v8/finance/chart/{code}.T"
           f"?range={rng}&interval=1d&events=div%2Csplit")
    raw = _get(url)
    if not raw:
        return None, None
    try:
        res = json.loads(raw)["chart"]["result"][0]
    except (KeyError, IndexError, ValueError, TypeError):
        return None, None

    ts = res.get("timestamp") or []
    q = res["indicators"]["quote"][0]
    # adjclose があれば分割・配当調整済みの終値比率で OHLC を揃える
    adj = None
    try:
        adj = res["indicators"]["adjclose"][0]["adjclose"]
    except (KeyError, IndexError, TypeError):
        pass

    bars = []
    for i, t in enumerate(ts):
        o, h, l, c, v = (q["open"][i], q["high"][i], q["low"][i],
                         q["close"][i], q["volume"][i])
        if None in (o, h, l, c):
            continue
        ratio = 1.0
        if adj and i < len(adj) and adj[i] and c:
            ratio = adj[i] / c
        bars.append({
            "date": datetime.fromtimestamp(t, JST).strftime("%Y-%m-%d"),
            "open": o * ratio, "high": h * ratio, "low": l * ratio,
            "close": c * ratio, "volume": v or 0,
            "close_raw": c,
        })
    splits = (res.get("events") or {}).get("splits") or {}
    return bars, splits


def fetch_tdnet_list(code, limit=80):
    url = f"https://webapi.yanoshin.jp/webapi/tdnet/list/{code}.json?limit={limit}"
    raw = _get(url)
    if not raw:
        return []
    try:
        items = json.loads(raw).get("items", [])
    except ValueError:
        return []
    out = []
    for it in items:
        d = it.get("Tdnet") or {}
        out.append({
            "id": d.get("document_id", ""),
            "title": d.get("title", ""),
            "date": (d.get("pubdate") or "")[:10],
            "url": d.get("document_url", ""),
        })
    return out


def fetch_tdnet_pdf(doc_id, yyyymmdd):
    CACHE_DIR.mkdir(exist_ok=True)
    cached = CACHE_DIR / f"{doc_id}.pdf"
    if cached.exists():
        return cached
    for url in (f"https://tdnet-pdf.kabutan.jp/{yyyymmdd}/{doc_id}.pdf",
                f"https://www.release.tdnet.info/inbs/{doc_id}.pdf"):
        blob = _get(url, binary=True)
        if blob and blob[:4] == b"%PDF":
            cached.write_bytes(blob)
            return cached
        time.sleep(SLEEP)
    return None


def pdf_text(path, max_pages=3):
    """短信は先頭のサマリーページに必要な数字が載る。"""
    if pdfplumber is None or path is None:
        return ""
    try:
        with pdfplumber.open(path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages[:max_pages])
    except Exception:
        return ""


def fetch_margin(code):
    """
    信用残（週次）。買残・売残の実株数を新しい順で返す。
    ページ構造が変わると取れないので、取れない場合は空を返し UNSET を伝播させる。
    """
    url = f"https://finance.yahoo.co.jp/quote/{code}.T/margin"
    html = _get(url)
    if not html:
        return []
    text = re.sub(r"<[^>]+>", "\t", html)
    text = text.replace("&nbsp;", " ")
    rows = []
    # 「YYYY/MM/DD ... 売残 買残 (倍率)」の並びを拾う
    pat = re.compile(
        r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})[^\d\-]{1,80}?"
        r"([\d,]{3,})[^\d\-]{1,40}?([\d,]{3,})"
    )
    for m in pat.finditer(text):
        y, mo, d, a, b = m.groups()
        try:
            sell, buy = int(a.replace(",", "")), int(b.replace(",", ""))
        except ValueError:
            continue
        rows.append({"date": f"{y}-{int(mo):02d}-{int(d):02d}",
                     "sell": sell, "buy": buy})
    seen, uniq = set(), []
    for r in rows:
        if r["date"] in seen:
            continue
        seen.add(r["date"])
        uniq.append(r)
    uniq.sort(key=lambda r: r["date"], reverse=True)
    return uniq[:8]


# ─────────────────────────── 計算層 ───────────────────────────

def atr14(bars, period=14):
    """Wilder の ATR。"""
    if len(bars) < period + 1:
        return None
    trs = []
    for i in range(1, len(bars)):
        h, l, pc = bars[i]["high"], bars[i]["low"], bars[i - 1]["close"]
        trs.append(max(h - l, abs(h - pc), abs(l - pc)))
    atr = sum(trs[:period]) / period
    for tr in trs[period:]:
        atr = (atr * (period - 1) + tr) / period
    return atr


def hi_lo(bars, n):
    w = bars[-n:]
    return max(b["high"] for b in w), min(b["low"] for b in w)


def adv(bars, n=20):
    """20日平均出来高（株）と平均売買代金（円）。"""
    w = bars[-n:]
    vol = sum(b["volume"] for b in w) / len(w)
    val = sum(b["close"] * b["volume"] for b in w) / len(w)
    return vol, val


def support_bands(bars, atr, cfg, max_bands=3):
    """
    スイング安値をクラスタリングして支持帯を抽出する。
    帯は [クラスタ内安値の最小, 最大]。タッチ日付を根拠として保持する。
    """
    k = cfg["swing_window"]
    window = bars[-cfg["lookback_bars_for_support"]:]
    pivots = []
    for i in range(k, len(window) - k):
        lo = window[i]["low"]
        if all(lo <= window[j]["low"] for j in range(i - k, i + k + 1)):
            pivots.append((lo, window[i]["date"]))
    if not pivots:
        return []

    tol = max(
        (cfg["support_cluster_tolerance_pct"] / 100.0) * window[-1]["close"],
        cfg["support_cluster_tolerance_atr"] * (atr or 0),
    )
    pivots.sort(key=lambda p: p[0])
    clusters, cur = [], [pivots[0]]
    for p in pivots[1:]:
        if p[0] - cur[0][0] <= tol:
            cur.append(p)
        else:
            clusters.append(cur)
            cur = [p]
    clusters.append(cur)

    bands = []
    for c in clusters:
        lows = [x[0] for x in c]
        dates = sorted(x[1] for x in c)
        bands.append({
            "low": min(lows), "high": max(lows),
            "touches": len(c), "dates": dates,
            "last_touch": dates[-1],
        })
    # タッチ回数を主、直近性を従にして強い帯を上位に
    bands.sort(key=lambda b: (b["touches"], b["last_touch"]), reverse=True)
    return bands[:max_bands]


def progress_rate(q_profit, fy_forecast):
    if not q_profit or not fy_forecast:
        return None
    return q_profit / fy_forecast * 100.0


# ─────────── 短信からの数値抽出（候補を出して人が確認する前提） ───────────

NUM = r"(△?▲?-?[\d,]+)"


def _to_num(s):
    if s is None:
        return None
    s = s.strip()
    neg = s.startswith(("△", "▲", "-"))
    s = re.sub(r"[△▲\-,]", "", s)
    if not s.isdigit():
        return None
    v = int(s)
    return -v if neg else v


def extract_tanshin_numbers(text):
    """
    短信サマリーから経常利益・営業利益・売上高の実績と通期予想の候補を拾う。
    誤読を避けるため、必ずマッチした原文行も返して人間が確認できるようにする。
    """
    out = {"actual": {}, "forecast": {}, "evidence": []}
    if not text:
        return out
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    keys = [("経常利益", "ordinary"), ("営業利益", "operating"), ("売上高", "revenue")]

    in_forecast = False
    for ln in lines:
        if re.search(r"(業績予想|通期.*予想|next|見通し)", ln):
            in_forecast = True
        for jp, en in keys:
            if jp in ln:
                nums = [_to_num(x) for x in re.findall(NUM, ln)]
                nums = [n for n in nums if n is not None and abs(n) >= 10]
                if not nums:
                    continue
                bucket = "forecast" if in_forecast else "actual"
                if en not in out[bucket]:
                    out[bucket][en] = nums[0]
                    out["evidence"].append(f"[{bucket}:{jp}] {ln}")
    return out


def pick_docs(items, patterns):
    return [it for it in items
            if any(re.search(p, it["title"]) for p in patterns)]


# ─────────────────────────── 設計層 ───────────────────────────

def design(name_cfg, market, fw):
    """
    構造 → 撤退ライン → サイズ の順で設計する。
    サイズは常に最後、かつ撤退ラインから逆算される。
    """
    close = market["close"]
    atr = market["atr14"]
    bands = market["support_bands"]
    warnings, skip_conditions = [], []

    if not bands or atr is None:
        return {"status": "設計不能", "reason": "支持帯またはATR14が未取得", "warnings": warnings}

    # ── ステップ1: 構造。現値より下にある最も高い支持帯を基準にする ──
    below = [b for b in bands if b["high"] < close]
    if not below:
        return {"status": "設計不能",
                "reason": "現値より下に支持帯が検出できない（押し目待ち以外の設計は禁止）",
                "warnings": warnings}
    base = max(below, key=lambda b: b["high"])

    # ── ステップ2: 撤退ライン。必ず支持帯の外側（下）に置く ──
    buffer = fw["stop_buffer_atr"] * atr
    stop = base["low"] - buffer
    stop = math.floor(stop)                      # 呼値未満は切り捨てて外側に寄せる
    if stop >= base["low"]:                      # 同値・内側は禁止
        stop = math.floor(base["low"]) - 1

    # ── ステップ3: エントリー。撤退ラインより必ず上 ──
    entry_low = base["low"]
    entry_high = base["high"] + 0.25 * atr
    if name_cfg.get("no_chase"):
        # 急騰後は現値より上に条件を置かない
        cap = close - 1
        entry_high = min(entry_high, cap)
        entry_low = min(entry_low, entry_high - 1)
    entry_low, entry_high = math.floor(entry_low), math.floor(entry_high)

    if entry_low <= stop:
        return {"status": "設計不能",
                "reason": "エントリー下限が撤退ラインを下回る（買い指値は撤退ラインより上、の制約に違反）",
                "warnings": warnings}

    # ── ATR距離の妥当性。近すぎ／遠すぎを明示する ──
    dist_hi = entry_high - stop          # 最悪エントリーからの距離（サイズ算出に使う）
    dist_lo = entry_low - stop
    mult_hi = dist_hi / atr
    mult_lo = dist_lo / atr
    if mult_hi < fw["atr_stop_multiple_min"]:
        warnings.append(f"撤退ラインが近すぎる（最悪エントリーで{mult_hi:.2f}×ATR14 < {fw['atr_stop_multiple_min']}×）")
    if mult_hi > fw["atr_stop_multiple_max"]:
        warnings.append(f"撤退ラインが遠すぎる（最悪エントリーで{mult_hi:.2f}×ATR14 > {fw['atr_stop_multiple_max']}×）")

    # ── ステップ4: サイズ。撤退ラインから逆算する ──
    risk_per_share = dist_hi
    lot = fw["lot_size"]
    shares_by_risk = int(450000 // risk_per_share) if risk_per_share > 0 else 0
    shares_by_risk = (shares_by_risk // lot) * lot

    adv20 = market["adv_vol20"]
    cap_pct = fw["liquidity_cap_pct_of_adv20"]
    shares_by_liq = int(adv20 * cap_pct / 100) if adv20 else 0
    shares_by_liq = (shares_by_liq // lot) * lot

    shares = min(shares_by_risk, shares_by_liq) if shares_by_liq else shares_by_risk
    binding = "流動性上限" if shares_by_liq and shares_by_liq < shares_by_risk else "損失上限"

    max_loss = shares * risk_per_share
    notional = shares * entry_high

    # ── 見送り条件 ──
    skip_conditions.append(
        f"終値が撤退ライン{stop:,.0f}円を下回った時点でエントリー自体を見送る（支持帯が壊れているため）")
    if market["margin_buy"] not in (None, UNSET) and market["days_to_cover"]:
        skip_conditions.append(
            f"信用買残が20日平均出来高の{market['days_to_cover']:.1f}日分から更に増加した週が出たら見送る")
    skip_conditions.append(
        f"判断期限{market['decision_deadline']}までにエントリーレンジに到達しない場合は発注せず見送る")

    return {
        "status": "設計可",
        "base_band": base,
        "stop": stop,
        "stop_basis": f"支持帯 {base['low']:,.0f}〜{base['high']:,.0f}円（{'/'.join(base['dates'][-3:])}）の外側 -{buffer:,.0f}円",
        "entry_low": entry_low,
        "entry_high": entry_high,
        "atr_multiple_at_entry_high": mult_hi,
        "atr_multiple_at_entry_low": mult_lo,
        "risk_per_share": risk_per_share,
        "shares_by_risk": shares_by_risk,
        "shares_by_liquidity": shares_by_liq,
        "shares": shares,
        "binding_constraint": binding,
        "max_loss": max_loss,
        "notional": notional,
        "distance_to_trigger_pct": (close - entry_high) / close * 100.0,
        "warnings": warnings,
        "skip_conditions": skip_conditions,
    }


# ─────────────────────────── 実行 ───────────────────────────

def analyse(nc, fw):
    code = nc["code"]
    print(f"\n=== {nc['name']} ({code}) 取得中 ===", file=sys.stderr)

    bars, splits = fetch_bars(code)
    time.sleep(SLEEP)
    if not bars:
        return {"cfg": nc, "error": "株価データ取得失敗"}

    a = atr14(bars)
    hi5, lo5 = hi_lo(bars, 5)
    hi20, lo20 = hi_lo(bars, 20)
    vol20, val20 = adv(bars, 20)
    bands = support_bands(bars, a, fw)

    margin = fetch_margin(code)
    time.sleep(SLEEP)
    mbuy = margin[0]["buy"] if margin else None
    msell = margin[0]["sell"] if margin else None
    dtc = (mbuy / vol20) if (mbuy and vol20) else None

    tdnet = fetch_tdnet_list(code)
    time.sleep(SLEEP)

    close = bars[-1]["close"]
    hint = nc.get("next_earnings_hint", "")
    deadline = UNSET
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", hint or "")
    if m:
        d = datetime.strptime(hint, "%Y-%m-%d")
        deadline = (d - timedelta(days=fw["decision_deadline_days_before_earnings"])).strftime("%Y-%m-%d")

    market = {
        "close": close,
        "last_bar_date": bars[-1]["date"],
        "atr14": a,
        "hi5": hi5, "lo5": lo5, "hi20": hi20, "lo20": lo20,
        "adv_vol20": vol20, "adv_val20": val20,
        "support_bands": bands,
        "margin_rows": margin,
        "margin_buy": mbuy if mbuy is not None else UNSET,
        "margin_sell": msell if msell is not None else UNSET,
        "days_to_cover": dtc,
        "splits": splits,
        "next_earnings": hint or UNSET,
        "decision_deadline": deadline,
        "tdnet": tdnet[:15],
    }

    # 依頼者提供の支持帯があれば、検出結果より優先して混ぜる
    for kb in nc.get("known_support_bands", []):
        market["support_bands"].append({
            "low": kb["low"], "high": kb["high"], "touches": 0,
            "dates": [kb["basis"]], "last_touch": kb["basis"],
            "source": kb.get("source", "依頼者提供"),
        })

    d = design(nc, market, fw)
    return {"cfg": nc, "market": market, "design": d}


def fmt(v, unit="", nd=0):
    if v is None or v == UNSET:
        return UNSET
    if isinstance(v, float):
        return f"{v:,.{nd}f}{unit}"
    return f"{v:,}{unit}"


def report(results, fw):
    lines = []
    for r in results:
        nc = r["cfg"]
        lines.append(f"\n{'='*70}\n■ {nc['name']}({nc['code']})")
        if r.get("error"):
            lines.append(f"  {r['error']}")
            continue
        m, d = r["market"], r["design"]
        lines.append(f"  終値 {fmt(m['close'],'円')}  ({m['last_bar_date']})")
        lines.append(f"  ATR14 {fmt(m['atr14'],'円',1)}")
        lines.append(f"  直近5日 高{fmt(m['hi5'],'円')} / 安{fmt(m['lo5'],'円')}")
        lines.append(f"  直近20日 高{fmt(m['hi20'],'円')} / 安{fmt(m['lo20'],'円')}")
        lines.append(f"  20日平均出来高 {fmt(m['adv_vol20'],'株')} / 20日平均売買代金 {fmt(m['adv_val20']/1e6,'百万円',1)}")
        lines.append("  支持帯:")
        for b in m["support_bands"]:
            src = b.get("source", f"{b['touches']}タッチ")
            lines.append(f"    {fmt(b['low'],'')}〜{fmt(b['high'],'円')}  ({src}: {'/'.join(b['dates'][-3:])})")
        lines.append(f"  信用買残 {fmt(m['margin_buy'],'株')} / 売残 {fmt(m['margin_sell'],'株')}")
        if m["days_to_cover"]:
            lines.append(f"  買残÷20日平均出来高 = {m['days_to_cover']:.1f}日分")
        for row in m["margin_rows"][:4]:
            lines.append(f"    {row['date']}  買{row['buy']:,} / 売{row['sell']:,}")
        lines.append(f"  次回決算 {m['next_earnings']} / 判断期限(2週前) {m['decision_deadline']}")

        if d["status"] != "設計可":
            lines.append(f"  ▶ {d['status']}: {d['reason']}")
            continue
        lines.append(f"  ▶ エントリー {fmt(d['entry_low'])}〜{fmt(d['entry_high'],'円')} × {fmt(d['shares'],'株')}")
        lines.append(f"     サイズ根拠: 450,000 ÷ {fmt(d['risk_per_share'],'円',1)}(=最悪{fmt(d['entry_high'])}−撤退{fmt(d['stop'])}) "
                     f"= {fmt(d['shares_by_risk'],'株')} / 流動性上限 {fmt(d['shares_by_liquidity'],'株')} → 拘束条件={d['binding_constraint']}")
        lines.append(f"  ▶ 撤退ライン {fmt(d['stop'],'円')}  {d['stop_basis']}")
        lines.append(f"     ATR14の {d['atr_multiple_at_entry_high']:.2f}倍（レンジ下限からは {d['atr_multiple_at_entry_low']:.2f}倍）")
        lines.append(f"  ▶ 想定最大損失 {fmt(d['max_loss'],'円')} / 建玉 {fmt(d['notional'],'円')}")
        for w in d["warnings"]:
            lines.append(f"  ⚠ {w}")
        for s in d["skip_conditions"]:
            lines.append(f"  ✕ 見送り条件: {s}")

    # 最終表: トリガーまでの距離が近い順
    lines.append(f"\n{'='*70}\n■ 一覧（トリガーまでの距離が近い順）")
    hdr = f"{'銘柄':<14}{'終値':>9}{'エントリー':>18}{'株数':>9}{'撤退':>9}{'最大損失':>11}{'距離%':>8}"
    lines.append(hdr)
    rows = [r for r in results if r.get("design", {}).get("status") == "設計可"]
    rows.sort(key=lambda r: r["design"]["distance_to_trigger_pct"])
    for r in rows:
        d, nc = r["design"], r["cfg"]
        lines.append(
            f"{nc['name']+'('+nc['code']+')':<14}"
            f"{r['market']['close']:>9,.0f}"
            f"{str(int(d['entry_low']))+'-'+str(int(d['entry_high'])):>18}"
            f"{d['shares']:>9,}"
            f"{d['stop']:>9,.0f}"
            f"{d['max_loss']:>11,.0f}"
            f"{d['distance_to_trigger_pct']:>8.2f}")
    for r in results:
        if r.get("design", {}).get("status") != "設計可":
            lines.append(f"{r['cfg']['name']+'('+r['cfg']['code']+')':<14}  → {r.get('design',{}).get('status', r.get('error','?'))}")
    return "\n".join(lines)


def main():
    cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    fw = cfg["framework"]
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    names = [n for n in cfg["names"] if not args or n["code"] in args]

    results = [analyse(n, fw) for n in names]
    print(report(results, fw))

    if "--json" in sys.argv:
        out = SCRIPT_DIR / "output.json"
        out.write_text(json.dumps(results, ensure_ascii=False, indent=2, default=str),
                       encoding="utf-8")
        print(f"\nraw → {out}", file=sys.stderr)


if __name__ == "__main__":
    main()
