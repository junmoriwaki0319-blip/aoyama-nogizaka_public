#!/usr/bin/env python3
"""
商社セクター・ダッシュボード フェーズ1 — 17 KPI 比較マトリクス生成

入力:
  - data/sogo-shosha/raw/<ticker>/yahoo.json  (fetch_yahoo.js が生成)
  - data/edinet-financials.json                 (既存 BS キャッシュ。任意フォールバック)

出力:
  - data/sogo-shosha/processed/<ticker>-summary.json
  - data/sogo-shosha/processed/comparison-matrix.csv  (BOM 付 UTF-8)
  - data/sogo-shosha/processed/benchmarks.json

KPI と算出可能性 (詳細は data/sogo-shosha/known-issues.md):
  A. 株主還元
    1. DOE                           = dividendRate / bookValue (per share base)        ✓
    2. 総還元性向                    = (配当 + 自己株買い) / 純利益                    ✗ Yahoo に自己株買い実施額なし
    3. 自己株買い実施額 (12M累計)   ✗ Yahoo 非公開
    4. 配当性向                      = summaryDetail.payoutRatio                       ✓
    5. 政策保有株式 / 純資産比率    ✗ 有報・統合報告書スクレイピング必要
  B. 資本効率
    6. ROE                           = financialData.returnOnEquity                    ✓
    7. ROIC (簡易)                   = NI / (Equity + Debt)                            △ 簡易推計
    8. PBR                           = defaultKeyStatistics.priceToBook                ✓
    9. PER                           = summaryDetail.trailingPE                        ✓
   10. EV/EBITDA                    = defaultKeyStatistics.enterpriseToEbitda          ✓
  C. 事業ポートフォリオ
   11. 資源/非資源 営業利益比率    ✗ IR 説明会資料スクレイピング必要
   12. 海外売上高比率              ✗ 有報セグメント情報スクレイピング必要
   13. セグメント別 ROIC           ✗ 一部社のみ開示。非開示社多い
  D. 財務健全性
   14. ネット有利子負債/EBITDA     = (totalDebt - totalCash) / ebitda                  ✓
   15. 自己資本比率                ✗ Yahoo に総資産情報なし。EDINET 詳細パース必要
   16. インタレスト・カバレッジ    ✗ Yahoo に支払利息情報なし
  E. アクティビスト・シグナル
   17. 現預金 / 時価総額比率       = totalCash / marketCap                              ✓
   18. 大量保有報告件数 (12M)      ✗ EDINET API キー未取得
   19. 政策保有縮減率 (過去5期)    ✗ 統合報告書スクレイピング必要

注: 元タスク仕様の合計記載 ("17 KPI") と本実装 (19 KPI) には差がある。
    v0.2 設計ドラフト §4 の A-E カテゴリの個別項目を全て計上した結果。
    詳細は known-issues.md の「KPI 件数の差異」セクション参照。
"""

import json
import os
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]
RAW = BASE / "data" / "sogo-shosha" / "raw"
PROC = BASE / "data" / "sogo-shosha" / "processed"
EDINET_CACHE = BASE / "data" / "edinet-financials.json"

TARGETS = [
    {"ticker": "8058", "name": "三菱商事",   "tier": "big5"},
    {"ticker": "8031", "name": "三井物産",   "tier": "big5"},
    {"ticker": "8001", "name": "伊藤忠商事", "tier": "big5"},
    {"ticker": "8053", "name": "住友商事",   "tier": "big5"},
    {"ticker": "8002", "name": "丸紅",       "tier": "big5"},
    {"ticker": "2768", "name": "双日",       "tier": "mid"},
    {"ticker": "8015", "name": "豊田通商",   "tier": "mid"},
]

# 17(=19) KPI 定義: (id, label, category, value_fn, format)
# value_fn: dict(yahoo) -> float | None
# fmt: "%", "x", "ratio" (.2), "money_jpy", "missing"


def safe(d, *path):
    cur = d
    for p in path:
        if cur is None or not isinstance(cur, dict):
            return None
        cur = cur.get(p)
    return cur


def kpi_doe(y):
    """DOE = 1株配当 / 1株株主資本"""
    div = safe(y, "summaryDetail", "dividendRate")
    book = safe(y, "defaultKeyStatistics", "bookValue")
    if div is None or book is None or book == 0:
        return None
    return round(div / book * 100, 2)  # %


def kpi_payout(y):
    p = safe(y, "summaryDetail", "payoutRatio")
    if p is None:
        return None
    return round(p * 100, 2)


def kpi_roe(y):
    r = safe(y, "financialData", "returnOnEquity")
    if r is None:
        return None
    return round(r * 100, 2)


def kpi_roic_simple(y):
    """簡易 ROIC = 当期純利益 / (株主資本 + 有利子負債)"""
    ni = safe(y, "defaultKeyStatistics", "netIncomeToCommon")
    book = safe(y, "defaultKeyStatistics", "bookValue")
    shares = safe(y, "defaultKeyStatistics", "sharesOutstanding")
    debt = safe(y, "financialData", "totalDebt")
    if None in (ni, book, shares, debt):
        return None
    equity = book * shares
    ic = equity + debt
    if ic <= 0:
        return None
    return round(ni / ic * 100, 2)


def kpi_pbr(y):
    v = safe(y, "defaultKeyStatistics", "priceToBook")
    return round(v, 2) if v is not None else None


def kpi_per(y):
    v = safe(y, "summaryDetail", "trailingPE")
    return round(v, 2) if v is not None else None


def kpi_ev_ebitda(y):
    v = safe(y, "defaultKeyStatistics", "enterpriseToEbitda")
    return round(v, 2) if v is not None else None


def kpi_net_debt_ebitda(y):
    debt = safe(y, "financialData", "totalDebt")
    cash = safe(y, "financialData", "totalCash")
    ebitda = safe(y, "financialData", "ebitda")
    if None in (debt, cash, ebitda) or ebitda == 0:
        return None
    return round((debt - cash) / ebitda, 2)


def kpi_cash_to_mcap(y):
    cash = safe(y, "financialData", "totalCash")
    mcap = safe(y, "summaryDetail", "marketCap") or safe(y, "quote", "marketCap") or safe(y, "price", "marketCap")
    if not cash or not mcap:
        return None
    return round(cash / mcap * 100, 2)


KPI_DEFS = [
    # id, label, category, fn, unit
    ("doe",                "DOE (株主資本配当率)",         "A",  kpi_doe,             "%"),
    ("total_payout_ratio", "総還元性向",                   "A",  None,                "%"),
    ("buyback_amount",     "自己株買い (12M, 億円)",       "A",  None,                "億円"),
    ("payout_ratio",       "配当性向",                     "A",  kpi_payout,          "%"),
    ("crossheld_to_eq",    "政策保有 / 純資産",            "A",  None,                "%"),
    ("roe",                "ROE",                          "B",  kpi_roe,             "%"),
    ("roic",               "ROIC (簡易: NI/IC)",           "B",  kpi_roic_simple,     "%"),
    ("pbr",                "PBR",                          "B",  kpi_pbr,             "倍"),
    ("per",                "PER",                          "B",  kpi_per,             "倍"),
    ("ev_ebitda",          "EV/EBITDA",                    "B",  kpi_ev_ebitda,       "倍"),
    ("resource_ratio",     "資源 営業利益比率",            "C",  None,                "%"),
    ("overseas_revenue",   "海外売上高比率",               "C",  None,                "%"),
    ("segment_roic",       "セグメント別 ROIC 開示",       "C",  None,                "有/無"),
    ("net_debt_ebitda",    "ネット有利子負債/EBITDA",      "D",  kpi_net_debt_ebitda, "倍"),
    ("equity_ratio",       "自己資本比率",                 "D",  None,                "%"),
    ("interest_coverage",  "インタレスト・カバレッジ",     "D",  None,                "倍"),
    ("cash_to_mcap",       "現預金 / 時価総額",            "E",  kpi_cash_to_mcap,    "%"),
    ("large_holder_count", "大量保有報告件数 (12M)",       "E",  None,                "件"),
    ("crossheld_reduction","政策保有縮減率 (5期)",         "E",  None,                "%"),
]


def load_yahoo(ticker):
    p = RAW / ticker / "yahoo.json"
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


def build_summary(t, yahoo):
    summary = {
        "ticker": t["ticker"],
        "name": t["name"],
        "tier": t["tier"],
        "fetchedAt": yahoo.get("fetchedAt") if yahoo else None,
        "currency": safe(yahoo, "summaryDetail", "currency"),
        "marketCap_JPY": safe(yahoo, "summaryDetail", "marketCap")
                         or safe(yahoo, "quote", "marketCap"),
        "lastFiscalYearEnd": safe(yahoo, "defaultKeyStatistics", "lastFiscalYearEnd"),
        "kpis": {},
    }
    for kid, label, cat, fn, unit in KPI_DEFS:
        val = fn(yahoo) if (fn is not None and yahoo is not None) else None
        summary["kpis"][kid] = {"label": label, "category": cat, "unit": unit, "value": val}
    return summary


def fmt_cell(val, unit):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        if unit in ("%",):
            return f"{val:.2f}"
        if unit in ("倍",):
            return f"{val:.2f}"
        if unit in ("億円",):
            return f"{val:.0f}"
        return f"{val}"
    return str(val)


def write_csv(summaries, path):
    """7 社 × 19 KPI のマトリクス CSV を BOM 付 UTF-8 で出力"""
    rows = []
    header = ["category", "kpi_id", "kpi_label", "unit"] + [s["ticker"] for s in summaries]
    rows.append(header)

    for kid, label, cat, _, unit in KPI_DEFS:
        row = [cat, kid, label, unit]
        for s in summaries:
            v = s["kpis"][kid]["value"]
            row.append(fmt_cell(v, unit))
        rows.append(row)

    # BOM 付 UTF-8
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        for r in rows:
            f.write(",".join('"' + str(c).replace('"', '""') + '"' for c in r) + "\r\n")


def write_md(summaries, path):
    """comparison-matrix.csv のマークダウン版"""
    lines = []
    lines.append("# 商社7社 比較マトリクス\n")
    lines.append(f"取得日: {datetime.now().strftime('%Y-%m-%d')}\n")
    lines.append("")
    lines.append("空欄 = データ取得不能。詳細は `known-issues.md` 参照。")
    lines.append("")

    # ヘッダ
    name_row = "| カテゴリ | KPI | 単位 | " + " | ".join(s["ticker"] + " " + s["name"] for s in summaries) + " |"
    sep_row  = "|----------|-----|------|" + "|".join("-" * 12 for _ in summaries) + "|"
    lines.append(name_row)
    lines.append(sep_row)

    for kid, label, cat, _, unit in KPI_DEFS:
        row = [cat, label, unit]
        for s in summaries:
            v = s["kpis"][kid]["value"]
            row.append(fmt_cell(v, unit))
        lines.append("| " + " | ".join(row) + " |")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_benchmarks(summaries):
    """5大商社平均と7社平均"""
    big5 = [s for s in summaries if s["tier"] == "big5"]
    all7 = summaries
    bench = {
        "as_of": datetime.now().strftime("%Y-%m-%d"),
        "big5": {"members": [s["ticker"] for s in big5], "kpis": {}},
        "all7": {"members": [s["ticker"] for s in all7], "kpis": {}},
    }
    for kid, _, _, _, _ in KPI_DEFS:
        for group_name, group in [("big5", big5), ("all7", all7)]:
            vals = [s["kpis"][kid]["value"] for s in group if s["kpis"][kid]["value"] is not None]
            if not vals:
                bench[group_name]["kpis"][kid] = {"avg": None, "n": 0}
            else:
                bench[group_name]["kpis"][kid] = {
                    "avg": round(sum(vals) / len(vals), 3),
                    "min": round(min(vals), 3),
                    "max": round(max(vals), 3),
                    "n": len(vals),
                }
    return bench


def main():
    PROC.mkdir(parents=True, exist_ok=True)
    print(f"=== build_matrix.py: 17(=19) KPI マトリクス生成 ===")
    print(f"入力: {RAW}")
    print(f"出力: {PROC}\n")

    summaries = []
    for t in TARGETS:
        yahoo = load_yahoo(t["ticker"])
        if yahoo is None:
            print(f"  WARN: {t['ticker']} {t['name']}: yahoo.json なし")
            summaries.append(build_summary(t, None))
            continue
        s = build_summary(t, yahoo)
        summaries.append(s)
        out = PROC / f"{t['ticker']}-summary.json"
        out.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")
        filled = sum(1 for k in s["kpis"].values() if k["value"] is not None)
        print(f"  {t['ticker']} {t['name']}: {filled}/{len(KPI_DEFS)} KPI")

    # CSV (BOM)
    csv_path = PROC / "comparison-matrix.csv"
    write_csv(summaries, csv_path)
    print(f"\n  → {csv_path}")

    # MD
    md_path = BASE / "data" / "sogo-shosha" / "comparison-matrix.md"
    write_md(summaries, md_path)
    print(f"  → {md_path}")

    # ベンチマーク
    bench = build_benchmarks(summaries)
    bench_path = PROC / "benchmarks.json"
    bench_path.write_text(json.dumps(bench, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  → {bench_path}")

    # 集計
    total_cells = len(KPI_DEFS) * len(summaries)
    filled_cells = sum(
        1 for s in summaries for k in s["kpis"].values() if k["value"] is not None
    )
    print(f"\n=== 完了 ===")
    print(f"  埋まり率: {filled_cells}/{total_cells} ({filled_cells / total_cells * 100:.1f}%)")


if __name__ == "__main__":
    main()
