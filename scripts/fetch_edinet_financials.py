#!/usr/bin/env python3
"""
EDINET XBRL 財務データキャッシュ生成スクリプト
主要銘柄の有価証券報告書からBS・含み益・土地データを抽出しJSONキャッシュに保存
"""

import json
import os
import re
import sys
import time
import zipfile
import io
from datetime import datetime, timedelta
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

API_KEY = os.environ.get("EDINET_API_KEY", "")
EDINET_API = "https://api.edinet-fsa.go.jp/api/v2"
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "edinet-financials.json")

# 全プリセット銘柄コード（重複排除）
TARGET_CODES = sorted(set([
    # メガバンク・大手金融
    '8306', '8316', '8411', '8308', '8309', '8601', '8604', '8630', '8725', '8766',
    # 低PBR注目
    '5411', '5401', '7011', '7012', '6301', '6302', '5020', '5019', '5021', '8058',
    '8053', '3402', '3401', '4183', '4188', '5332', '5333', '3861', '3863', '4042',
    # 自動車
    '7203', '7267', '7201', '7270', '7269', '7211', '7202', '7272', '5108', '6902',
    # 総合商社
    '8031', '8001', '8002', '8015', '8020',
    # IT・通信
    '9984', '9433', '9432', '9434', '4689', '4755', '3382', '6758', '6861', '6954',
    # コングロマリット
    '6501', '6502', '6503', '7751', '6752', '6753', '6702', '6701', '8035', '7731',
    # 不動産
    '8801', '8802', '8804', '8830', '3289', '3291', '3462', '8818', '8848', '3003',
    # 食品・飲料
    '2502', '2503', '2587', '2801', '2802', '2871', '2897', '2914', '2269', '2809',
    # 医薬品
    '4502', '4503', '4506', '4507', '4519', '4523', '4568', '4578', '4151', '4452',
    # 素材・化学
    '4063', '4005', '4004', '4021', '4208', '4631', '4901',
    # 鉄鋼
    '5406', '5423', '5444', '5471', '5480', '5481', '5486', '5801',
    # 建設
    '1801', '1802', '1803', '1812', '1820', '1821', '1878', '1925', '1928', '1963',
    # 小売
    '8267', '8252', '9983', '7532', '3086', '3099', '2651', '7453', '2670',
    # 鉄道・運輸
    '9020', '9021', '9022', '9001', '9005', '9007', '9064', '9101', '9104', '9107',
    # 電力・ガス
    '9501', '9502', '9503', '9504', '9505', '9506', '9507', '9508', '9509', '9531',
    # 機械
    '6305', '6361', '6367', '6471', '6473', '7004',
    # 電機・精密
    '6971', '6762', '7741',
]))


def fetch_json(url, retries=3):
    """EDINET APIからJSONを取得"""
    for attempt in range(retries):
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urlopen(req, timeout=15) as resp:
                return json.loads(resp.read())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
            else:
                print(f"  WARN: fetch failed: {url[:80]}... - {e}")
                return None


def fetch_binary(url, retries=3):
    """バイナリデータを取得"""
    for attempt in range(retries):
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urlopen(req, timeout=60) as resp:
                return resp.read()
        except (URLError, HTTPError) as e:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
            else:
                print(f"  WARN: download failed: {url[:80]}... - {e}")
                return None


def search_annual_reports():
    """過去400日分のEDINET書類リストから有価証券報告書を検索"""
    print("=== EDINET書類検索 ===")
    sec_code_to_doc = {}  # secCode(4桁) → 最新docInfo
    today = datetime.now()
    dates = []
    for i in range(400):
        d = today - timedelta(days=i)
        if d.weekday() < 5:  # 平日のみ
            dates.append(d.strftime("%Y-%m-%d"))

    # 20日ずつバッチで検索
    BATCH = 20
    for batch_idx in range(0, len(dates), BATCH):
        batch_dates = dates[batch_idx:batch_idx + BATCH]
        for date_str in batch_dates:
            url = f"{EDINET_API}/documents.json?date={date_str}&type=2&Subscription-Key={API_KEY}"
            data = fetch_json(url)
            if not data or "results" not in data:
                continue
            for doc in data["results"]:
                if doc.get("docTypeCode") != "120":
                    continue
                raw_sec = doc.get("secCode") or ""
                sec4 = raw_sec[:4] if len(raw_sec) >= 4 else ""
                if not sec4 or sec4 not in TARGET_CODES:
                    continue
                # 最新のものを保持
                if sec4 not in sec_code_to_doc:
                    sec_code_to_doc[sec4] = {
                        "docID": doc["docID"],
                        "filerName": doc.get("filerName", ""),
                        "periodEnd": doc.get("periodEnd", ""),
                        "submitDateTime": doc.get("submitDateTime", ""),
                    }
                    print(f"  Found: {sec4} {doc.get('filerName','')} ({doc['docID']})")

        # 全ターゲット発見したら早期終了
        if len(sec_code_to_doc) >= len(TARGET_CODES):
            break
        time.sleep(0.5)

        progress = len(sec_code_to_doc)
        searched = min(batch_idx + BATCH, len(dates))
        print(f"  ... searched {searched}/{len(dates)} dates, found {progress}/{len(TARGET_CODES)} companies")

    print(f"\n合計 {len(sec_code_to_doc)} 社の有価証券報告書を発見")
    return sec_code_to_doc


def download_and_parse_xbrl(doc_id):
    """XBRL ZIPをダウンロードして財務データを抽出"""
    url = f"{EDINET_API}/documents/{doc_id}?type=1&Subscription-Key={API_KEY}"
    zip_data = fetch_binary(url)
    if not zip_data:
        return None

    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_data))
    except zipfile.BadZipFile:
        print(f"  WARN: Bad ZIP for {doc_id}")
        return None

    # XBRLインスタンスドキュメントを探す
    xbrl_name = None
    for name in zf.namelist():
        if name.endswith(".xbrl") and "AuditDoc" not in name and "PublicDoc" in name:
            xbrl_name = name
            break
    if not xbrl_name:
        for name in zf.namelist():
            if name.endswith(".xbrl") and "AuditDoc" not in name:
                xbrl_name = name
                break
    if not xbrl_name:
        return None

    xbrl_xml = zf.read(xbrl_name).decode("utf-8", errors="replace")
    result = parse_xbrl(xbrl_xml)

    # 注記HTMLから投資不動産・有価証券データを抽出
    for name in zf.namelist():
        if re.search(r"\.htm(l)?$", name, re.I) and "PublicDoc" in name:
            try:
                html = zf.read(name).decode("utf-8", errors="replace")
                parse_notes(html, result)
            except Exception:
                pass

    # 土地含み益推定
    estimate_land_gain(result)

    zf.close()
    return result


def find_xbrl_value(xml, element_name, type_hint="Instant"):
    """XBRLから指定要素の値を取得（連結・当期を優先）"""
    pattern = re.compile(
        rf"<[^>]*?:?{element_name}[^>]*contextRef=\"([^\"]*?)\"[^>]*>([^<]+)<",
        re.IGNORECASE
    )
    best_val = None
    best_priority = -1

    for m in pattern.finditer(xml):
        ctx = m.group(1)
        raw = m.group(2).replace(",", "").strip()
        try:
            num = float(raw)
        except ValueError:
            continue

        priority = 0
        if type_hint and type_hint.lower() in ctx.lower():
            priority += 10
        if re.search(r"CurrentYear", ctx, re.I) or (re.search(r"Current", ctx, re.I) and not re.search(r"Prior", ctx, re.I)):
            priority += 5
        if not re.search(r"NonConsolidated", ctx, re.I):
            priority += 3
        if not re.search(r"Member", ctx, re.I) or re.search(r"ConsolidatedMember", ctx, re.I):
            priority += 1

        if priority > best_priority:
            best_priority = priority
            best_val = num

    return best_val


def parse_xbrl(xml):
    """XBRLからBS項目を抽出"""
    result = {}

    bs_items = {
        "cashAndDeposits": ["CashAndDeposits"],
        "shortTermSecurities": ["ShortTermInvestmentSecurities"],
        "investmentSecurities": ["InvestmentSecurities"],
        "land": ["Land"],
        "netAssets": ["NetAssets", "TotalNetAssets"],
        "shareholdersEquity": ["ShareholdersEquity", "TotalShareholdersEquity"],
        "shortTermBorrowings": ["ShortTermBorrowings", "ShortTermLoansPayable"],
        "currentPortionLongTermDebt": ["CurrentPortionOfLongTermLoansPayable", "CurrentPortionOfLongTermDebt"],
        "currentPortionBonds": ["CurrentPortionOfBondsPayable", "CurrentPortionOfBonds"],
        "longTermBorrowings": ["LongTermLoansPayable", "LongTermBorrowings"],
        "bondsPayable": ["BondsPayable"],
        "landRevaluationReserve": ["RevaluationReserveForLand", "LandRevaluationExcess"],
    }

    for key, elements in bs_items.items():
        for el in elements:
            val = find_xbrl_value(xml, el, "Instant")
            if val is not None:
                result[key] = round(val / 1_000_000)  # 円→百万円
                break

    # PL items (Duration)
    for el in ["OperatingIncome", "OperatingProfit"]:
        val = find_xbrl_value(xml, el, "Duration")
        if val is not None:
            result["operatingIncome"] = round(val / 1_000_000)
            break

    da_total = find_xbrl_value(xml, "DepreciationAndAmortization", "Duration")
    if da_total is None:
        da_total = find_xbrl_value(xml, "Depreciation", "Duration")
    if da_total is not None:
        result["depreciationAndAmortization"] = round(da_total / 1_000_000)
    else:
        da_sum = 0
        found = False
        for el in ["DepreciationCostOfSales", "DepreciationSGA"]:
            v = find_xbrl_value(xml, el, "Duration")
            if v is not None:
                da_sum += v
                found = True
        if found:
            result["depreciationAndAmortization"] = round(da_sum / 1_000_000)

    return result


def parse_notes(html, data):
    """注記テキストから投資不動産・有価証券データを抽出"""
    # 投資不動産
    if "investmentPropertyBookValue" not in data:
        ip_idx = html.find("投資不動産")
        if ip_idx > -1:
            search = html[max(0, ip_idx - 500):min(len(html), ip_idx + 5000)]
            bs_m = re.search(r"貸借対照表計上額[\s\S]{0,200}?([\d,]+)", search)
            fv_m = re.search(r"(?:時価|公正価値)[\s\S]{0,200}?([\d,]+)", search)
            if bs_m and fv_m:
                bv = int(bs_m.group(1).replace(",", ""))
                fv = int(fv_m.group(1).replace(",", ""))
                if bv > 0 and fv > 0:
                    data["investmentPropertyBookValue"] = bv
                    data["investmentPropertyFairValue"] = fv

    # 有価証券
    if "securitiesBookValue" not in data:
        sec_idx = html.find("その他有価証券")
        if sec_idx > -1:
            search = html[max(0, sec_idx - 200):min(len(html), sec_idx + 5000)]
            cost_m = re.search(r"取得原価[\s\S]{0,300}?([\d,]+)", search)
            bv_m = re.search(r"貸借対照表計上額[\s\S]{0,300}?([\d,]+)", search)
            if cost_m and bv_m:
                cost = int(cost_m.group(1).replace(",", ""))
                bv = int(bv_m.group(1).replace(",", ""))
                if cost > 0 and bv > 0:
                    data["securitiesBookValue"] = cost
                    data["securitiesMarketValue"] = bv


def estimate_land_gain(data):
    """土地含み益の推定"""
    data["estimatedLandGain"] = None
    data["landGainMethod"] = None

    land_bv = data.get("land")
    if not land_bv or land_bv <= 0:
        return

    # 方法1: 投資不動産比率準用
    ip_bv = data.get("investmentPropertyBookValue", 0)
    ip_fv = data.get("investmentPropertyFairValue", 0)
    if ip_bv > 0 and ip_fv > 0:
        ratio = ip_fv / ip_bv
        conservative = 1 + (ratio - 1) * 0.7
        data["estimatedLandGain"] = round(land_bv * conservative - land_bv)
        data["landGainMethod"] = f"投資不動産比率準用（時価/簿価={ratio:.2f}倍→保守的70%適用）"
        return

    # 方法2: 土地再評価差額金
    reserve = data.get("landRevaluationReserve", 0)
    if reserve and reserve != 0:
        data["estimatedLandGain"] = round(reserve / 0.7)
        data["landGainMethod"] = "土地再評価差額金から逆算（税効果30%戻し）"
        return

    data["landGainMethod"] = "推定データ不足"


def main():
    if not API_KEY:
        print("ERROR: EDINET_API_KEY environment variable not set")
        sys.exit(1)

    print(f"対象銘柄: {len(TARGET_CODES)} 社")
    print(f"出力先: {OUTPUT_PATH}\n")

    # 既存キャッシュを読み込み（差分更新のため）
    existing = {}
    if os.path.exists(OUTPUT_PATH):
        try:
            with open(OUTPUT_PATH, "r", encoding="utf-8") as f:
                existing = json.load(f).get("companies", {})
        except Exception:
            pass

    # Step 1: 有価証券報告書を検索
    doc_map = search_annual_reports()

    # Step 2: XBRLダウンロード & 解析
    print("\n=== XBRL解析 ===")
    companies = dict(existing)  # 既存データを維持
    processed = 0

    for code, doc_info in doc_map.items():
        # 既にキャッシュ済みで同じdocIDなら スキップ
        if code in companies and companies[code].get("_docID") == doc_info["docID"]:
            print(f"  SKIP: {code} {doc_info['filerName']} (cached)")
            continue

        print(f"  Processing: {code} {doc_info['filerName']}...", end=" ", flush=True)
        data = download_and_parse_xbrl(doc_info["docID"])
        if data:
            data["_docID"] = doc_info["docID"]
            data["_filerName"] = doc_info["filerName"]
            data["_periodEnd"] = doc_info["periodEnd"]
            data["_submitDateTime"] = doc_info["submitDateTime"]
            companies[code] = data
            items = sum(1 for v in data.values() if v is not None and not str(v).startswith("_"))
            print(f"OK ({items} items)")
        else:
            print("FAILED")

        processed += 1
        if processed % 5 == 0:
            time.sleep(1)  # EDINET API負荷軽減

    # Step 3: JSON保存
    output = {
        "last_updated": datetime.now().strftime("%Y-%m-%dT%H:%M:%S+09:00"),
        "total_companies": len(companies),
        "companies": companies,
    }
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n=== 完了: {len(companies)} 社のデータを保存 ===")


if __name__ == "__main__":
    main()
