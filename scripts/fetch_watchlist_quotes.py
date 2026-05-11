"""
Fetch quote / fundamental / margin / earnings-calendar data for a Japanese
equity watchlist from kabutan.jp (HTML scrape) and yfinance, then print a
Markdown table sorted by next earnings date.

Run locally (this sandbox blocks kabutan.jp and Yahoo Finance hosts):

    pip install yfinance beautifulsoup4 lxml requests
    python3 scripts/fetch_watchlist_quotes.py
"""

from __future__ import annotations

import dataclasses
import re
import sys
import time
from datetime import date, datetime
from typing import Optional

import requests
from bs4 import BeautifulSoup

try:
    import yfinance as yf
except ImportError:
    yf = None


WATCHLIST: list[tuple[str, str, str]] = [
    # (code, name, bucket)
    ("563A", "GXナスダック100カバコ", "明日発注"),
    ("2243", "GX半導体関連-日本株式", "明日発注"),
    ("3844", "コムチュア", "明日発注"),
    ("6584", "三桜工業", "明日発注"),
    ("4022", "ラサ工業", "ウォッチ"),
    ("7220", "武蔵精密工業", "ウォッチ"),
    ("6226", "守谷輸送機工業", "ウォッチ"),
    ("6144", "西部電機", "ウォッチ"),
    ("6890", "フェローテック", "ウォッチ"),
    ("1969", "高砂熱学工業", "ウォッチ"),
    ("1980", "ダイダン", "ウォッチ"),
    ("6490", "日本ピラー工業", "ウォッチ"),
    ("5217", "テクノクオーツ", "ウォッチ"),
]

UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://kabutan.jp/",
}


@dataclasses.dataclass
class Quote:
    code: str
    name: str
    bucket: str
    price: Optional[float] = None
    pts: Optional[float] = None
    per: Optional[float] = None
    pbr: Optional[float] = None
    market_cap_oku: Optional[float] = None  # 億円
    div_yield: Optional[float] = None       # %
    margin_buy: Optional[int] = None
    margin_sell: Optional[int] = None
    margin_ratio: Optional[float] = None
    earnings_date: Optional[str] = None
    op_progress_q3: Optional[str] = None
    nav: Optional[float] = None             # ETF基準価額
    nav_disparity: Optional[float] = None   # 乖離率 %
    last_distribution: Optional[str] = None
    notes: list[str] = dataclasses.field(default_factory=list)

    def fmt(self, attr: str, suffix: str = "", digits: int = 2) -> str:
        v = getattr(self, attr)
        if v is None or v == "":
            return "要確認"
        if isinstance(v, float):
            return f"{v:,.{digits}f}{suffix}"
        if isinstance(v, int):
            return f"{v:,}{suffix}"
        return str(v)


def _to_float(text: str) -> Optional[float]:
    if text is None:
        return None
    t = text.strip().replace(",", "").replace("円", "").replace("倍", "").replace("％", "").replace("%", "")
    if t in ("", "-", "－", "---"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def fetch_kabutan(q: Quote) -> None:
    url = f"https://kabutan.jp/stock/?code={q.code}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
    except Exception as e:
        q.notes.append(f"kabutan取得失敗: {e}")
        return

    soup = BeautifulSoup(r.text, "lxml")

    price_el = soup.select_one("span.kabuka") or soup.select_one("div.stock_kabuka_table strong")
    if price_el:
        q.price = _to_float(price_el.get_text())

    text = soup.get_text("\n")

    m = re.search(r"PER[^\n]*?連[^\n]*?([\d,.\-]+)\s*倍", text)
    if m:
        q.per = _to_float(m.group(1))
    m = re.search(r"PBR[^\n]*?([\d,.\-]+)\s*倍", text)
    if m:
        q.pbr = _to_float(m.group(1))
    m = re.search(r"利回り[^\n]*?([\d,.\-]+)\s*[％%]", text)
    if m:
        q.div_yield = _to_float(m.group(1))
    m = re.search(r"時価総額[^\n]*?([\d,]+)\s*百万円", text)
    if m:
        v = _to_float(m.group(1))
        if v is not None:
            q.market_cap_oku = v / 100.0  # 百万円→億円

    # 信用残（kabutan の信用残テーブルは /stock/?code=XXXX で別ページにあるためサブ取得）
    try:
        rr = requests.get(
            f"https://kabutan.jp/stock/kabuka?code={q.code}", headers=HEADERS, timeout=20
        )
        rr.raise_for_status()
        sub = BeautifulSoup(rr.text, "lxml")
        sub_text = sub.get_text("\n")
        m = re.search(r"信用買残[^\n]*?([\d,]+)", sub_text)
        if m:
            q.margin_buy = int(m.group(1).replace(",", ""))
        m = re.search(r"信用売残[^\n]*?([\d,]+)", sub_text)
        if m:
            q.margin_sell = int(m.group(1).replace(",", ""))
        m = re.search(r"信用倍率[^\n]*?([\d,.\-]+)", sub_text)
        if m:
            q.margin_ratio = _to_float(m.group(1))
    except Exception as e:
        q.notes.append(f"信用残取得失敗: {e}")

    # 決算発表予定日（kabutan の銘柄ページに「次回決算発表日」ブロックあり）
    m = re.search(r"次回[^\n]{0,10}決算[^\n]{0,10}(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})", text)
    if m:
        y, mo, d = (int(x) for x in m.groups())
        q.earnings_date = f"{y:04d}-{mo:02d}-{d:02d}"

    time.sleep(0.8)  # be polite


def fetch_yfinance(q: Quote) -> None:
    if yf is None:
        q.notes.append("yfinance未インストール")
        return
    sym = f"{q.code}.T"
    try:
        t = yf.Ticker(sym)
        info = t.info or {}
        if q.price is None and info.get("regularMarketPrice"):
            q.price = float(info["regularMarketPrice"])
        if q.per is None and info.get("forwardPE"):
            q.per = float(info["forwardPE"])
        if q.pbr is None and info.get("priceToBook"):
            q.pbr = float(info["priceToBook"])
        if q.div_yield is None and info.get("dividendYield"):
            dy = float(info["dividendYield"])
            q.div_yield = dy * 100 if dy < 1 else dy
        if q.market_cap_oku is None and info.get("marketCap"):
            q.market_cap_oku = float(info["marketCap"]) / 1e8  # 円→億円
        if q.earnings_date is None:
            cal = getattr(t, "calendar", None)
            if cal is not None and not getattr(cal, "empty", True):
                try:
                    ed = cal.loc["Earnings Date"]
                    if hasattr(ed, "iloc"):
                        ed = ed.iloc[0]
                    if isinstance(ed, (datetime, date)):
                        q.earnings_date = ed.strftime("%Y-%m-%d")
                except Exception:
                    pass
    except Exception as e:
        q.notes.append(f"yfinance取得失敗: {e}")


def render_markdown(rows: list[Quote]) -> str:
    def keyfn(q: Quote) -> tuple[int, str]:
        return (0, q.earnings_date) if q.earnings_date else (1, q.code)

    rows = sorted(rows, key=keyfn)

    header = (
        "| 決算日 | コード | 銘柄 | 区分 | 現在値 | PER | PBR | 時価総額(億) | "
        "信用買残 | 信用売残 | 信用倍率 | 配当利回り | Q3営利進捗 | 備考 |"
    )
    sep = "|" + "|".join(["---"] * 14) + "|"
    lines = [header, sep]
    for q in rows:
        lines.append(
            "| {ed} | {code} | {name} | {bucket} | {px} | {per} | {pbr} | {mc} | "
            "{mb} | {ms} | {mr} | {dy} | {prog} | {note} |".format(
                ed=q.earnings_date or "要確認",
                code=q.code,
                name=q.name,
                bucket=q.bucket,
                px=q.fmt("price"),
                per=q.fmt("per", "倍"),
                pbr=q.fmt("pbr", "倍"),
                mc=q.fmt("market_cap_oku", "", 1),
                mb=q.fmt("margin_buy"),
                ms=q.fmt("margin_sell"),
                mr=q.fmt("margin_ratio", "倍"),
                dy=q.fmt("div_yield", "%"),
                prog=q.op_progress_q3 or "要確認",
                note="; ".join(q.notes) if q.notes else "",
            )
        )
    return "\n".join(lines)


def main() -> int:
    rows: list[Quote] = []
    for code, name, bucket in WATCHLIST:
        q = Quote(code=code, name=name, bucket=bucket)
        fetch_kabutan(q)
        fetch_yfinance(q)
        rows.append(q)
        print(f"  fetched {code} {name}", file=sys.stderr)

    print(render_markdown(rows))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
