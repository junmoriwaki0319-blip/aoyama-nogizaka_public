#!/usr/bin/env python3
"""
EDINET API から大量保有報告書を取得し、JSON データを生成するスクリプト。
GitHub Actions で毎日 8:00 / 16:00 / 19:00 (JST) に実行される。
"""

import json
import os
import re
import sys
import time
import zipfile
import io
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlencode

import csv
import requests

# ─── 設定 ───
EDINET_API_BASE = "https://api.edinet-fsa.go.jp/api/v2"
API_KEY = os.environ.get("EDINET_API_KEY", "")
JST = timezone(timedelta(hours=9))
LOOKBACK_DAYS = 540  # 初回取得: 過去何日分を取得するか（約18ヶ月）
INCREMENTAL_DAYS = 7  # 増分取得: 既存データがある場合は直近何日分のみ取得
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR.parent / "data"
OUTPUT_FILE = DATA_DIR / "reports.json"
ACTIVISTS_FILE = SCRIPT_DIR / "known_activists.json"

# 大量保有報告書の docTypeCode (EDINET API v2)
DOC_TYPE_LARGE_HOLDING = "350"       # 大量保有報告書・変更報告書
DOC_TYPE_CORRECTION = "360"          # 訂正報告書（大量保有報告書・変更報告書）


def load_edinet_code_map():
    """EDINET コードリストをダウンロードし、edinetCode → 企業名のマッピングを構築"""
    if not API_KEY:
        return {}

    url = f"{EDINET_API_BASE}/EdinetcodeDlInfo/codes?Subscription-Key={API_KEY}"

    try:
        resp = requests.get(url, timeout=60)
        if resp.status_code != 200:
            print(f"  [WARN] EDINET code list download failed: {resp.status_code}", file=sys.stderr)
            return {}

        content_type = resp.headers.get("Content-Type", "")
        if "zip" not in content_type and "octet" not in content_type:
            print(f"  [WARN] Unexpected content-type for code list: {content_type}", file=sys.stderr)
            return {}

        code_map = {}
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            for name in zf.namelist():
                if name.endswith(".csv"):
                    raw = zf.read(name)
                    # Try UTF-8 with BOM, then shift_jis
                    for enc in ["utf-8-sig", "shift_jis", "cp932"]:
                        try:
                            text = raw.decode(enc)
                            break
                        except (UnicodeDecodeError, LookupError):
                            continue
                    else:
                        continue

                    reader = csv.reader(text.splitlines())
                    header = None
                    name_col = None
                    code_col = None
                    for row in reader:
                        if header is None:
                            header = row
                            # Find columns dynamically
                            for idx, col in enumerate(header):
                                col_clean = col.strip().replace('\ufeff', '')
                                if 'EDINET' in col_clean and 'コード' in col_clean:
                                    code_col = idx
                                elif '提出者名' in col_clean or '発行者名' in col_clean or '名称' in col_clean:
                                    if name_col is None:
                                        name_col = idx
                            # Fallback: first col = code, 7th col = name
                            if code_col is None:
                                code_col = 0
                            if name_col is None:
                                name_col = 6
                            continue
                        if len(row) <= max(code_col, name_col):
                            continue
                        edinet_code = row[code_col].strip()
                        company_name = row[name_col].strip()
                        if edinet_code and company_name:
                            code_map[edinet_code] = company_name
                    break  # Only need the first CSV

        print(f"EDINET コードリスト: {len(code_map)} 件")
        return code_map

    except Exception as e:
        print(f"  [WARN] EDINET code list error: {e}", file=sys.stderr)
        return {}


def load_known_activists():
    """既知のアクティビスト投資家リストを読み込む"""
    with open(ACTIVISTS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["activists"]


def match_activist(filer_name, activists):
    """報告者名が既知のアクティビストに該当するか判定"""
    for activist in activists:
        names_to_check = [activist["name"]] + activist.get("aliases", [])
        for name in names_to_check:
            if name in filer_name or filer_name in name:
                return activist
    return None


def fetch_document_list(date_str):
    """EDINET API から指定日の書類一覧を取得"""
    params = {
        "date": date_str,
        "type": 2,  # メタデータ + 提出書類一覧
        "Subscription-Key": API_KEY,
    }

    url = f"{EDINET_API_BASE}/documents.json?{urlencode(params)}"

    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return data.get("results", [])
    except requests.RequestException as e:
        print(f"  [WARN] API error for {date_str}: {e}", file=sys.stderr)
        return []


def filter_large_holdings(documents):
    """大量保有報告書 / 変更報告書 / 訂正報告書 のみ抽出"""
    target_types = {DOC_TYPE_LARGE_HOLDING, DOC_TYPE_CORRECTION}
    return [
        doc for doc in documents
        if doc.get("docTypeCode") in target_types
    ]


def extract_sec_code(raw_code):
    """5桁の証券コードから4桁に変換"""
    if not raw_code:
        return ""
    code = str(raw_code).strip()
    # 5桁なら末尾のチェックディジットを除去
    if len(code) == 5:
        return code[:4]
    return code


def download_xbrl_and_extract(doc_id):
    """
    XBRL ZIP をダウンロードし、保有比率・保有目的を抽出する。
    抽出できない場合は空の dict を返す。
    """
    if not API_KEY:
        return {}

    params = {"type": 1}  # XBRL ZIP
    if API_KEY:
        params["Subscription-Key"] = API_KEY

    url = f"{EDINET_API_BASE}/documents/{doc_id}?{urlencode(params)}"

    try:
        resp = requests.get(url, timeout=60)
        if resp.status_code != 200:
            return {}

        content_type = resp.headers.get("Content-Type", "")
        if "zip" not in content_type and "octet" not in content_type:
            return {}

        result = {}

        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            for name in zf.namelist():
                if name.endswith(".xbrl") or name.endswith(".htm") or name.endswith(".html"):
                    try:
                        text = zf.read(name).decode("utf-8", errors="ignore")
                    except Exception:
                        continue

                    # 発行者名（対象企業名）の抽出
                    issuer_patterns = [
                        # XBRL element for issuer name
                        r'name="[^"]*(?:[Ii]ssuer[Nn]ame|NameOfIssuer|IssuerNameJp)[^"]*"[^>]*>([^<]+)',
                        # 「発行者の名称」セクション
                        r'発行者の名称[^：:]*[：:]\s*([^\n<]{2,40}?)(?:\s*[（(]|$|\s{2})',
                        r'発行者の名称.*?<[^>]*>\s*([^\n<]{2,40}?)\s*<',
                        # 「株券等の発行者」セクション
                        r'株券等の発行者[^：:]*[：:]\s*([^\n<]{2,40}?)(?:\s*[（(]|$|\s{2})',
                    ]
                    for pattern in issuer_patterns:
                        m = re.search(pattern, text)
                        if m:
                            name = m.group(1).strip()
                            # Filter out generic text and filer's own name
                            if (name and len(name) >= 2
                                    and '報告書' not in name
                                    and '提出者' not in name
                                    and '代表取締役' not in name):
                                result["target_company"] = name
                                break

                    # 証券コードの抽出
                    code_patterns = [
                        r'(?:証券コード|銘柄コード)[^\d]{0,10}(\d{4,5})',
                        r'name="[^"]*(?:[Ss]ecurity[Cc]ode|SecuritiesCode)[^"]*"[^>]*>(\d{4,5})',
                    ]
                    for pattern in code_patterns:
                        m = re.search(pattern, text)
                        if m:
                            result["sec_code"] = m.group(1).strip()
                            break

                    # 保有割合の抽出（複数パターン対応）
                    ratio_patterns = [
                        r'(?:保有割合|所有割合)[^\d]{0,30}?([\d]+[\.．][\d]+)\s*[%％]',
                        r'name="[^"]*(?:HoldingRatio|OwnershipRatio)[^"]*"[^>]*>([\d]+[\.．][\d]+)',
                        r'([\d]+[\.．][\d]+)\s*[%％]\s*(?:（.*?保有割合|を保有)',
                    ]
                    for pattern in ratio_patterns:
                        m = re.search(pattern, text)
                        if m:
                            ratio_str = m.group(1).replace("．", ".")
                            try:
                                result["holding_ratio"] = float(ratio_str)
                            except ValueError:
                                pass
                            break

                    # 保有目的の抽出
                    purpose_patterns = [
                        r'保有目的[^\n]{0,5}[：:]\s*([^\n<]{2,60})',
                        r'(?:当該株券等の発行者の事業活動を|純投資|投資及び状況に応じて|政策投資|経営参加|株主提案)[^\n<]{0,80}',
                    ]
                    for pattern in purpose_patterns:
                        m = re.search(pattern, text)
                        if m:
                            purpose_raw = m.group(1) if m.lastindex else m.group(0)
                            purpose_raw = purpose_raw.strip()
                            result["purpose"] = classify_purpose(purpose_raw)
                            result["purpose_detail"] = purpose_raw[:100]
                            break

                    if result:
                        break

        return result

    except Exception as e:
        print(f"  [WARN] XBRL parse error for {doc_id}: {e}", file=sys.stderr)
        return {}


def classify_purpose(purpose_text):
    """保有目的テキストを分類"""
    if "純投資" in purpose_text:
        return "純投資"
    if "政策" in purpose_text:
        return "政策投資"
    if "株主提案" in purpose_text or "提案" in purpose_text:
        return "株主提案"
    if "経営" in purpose_text or "支配" in purpose_text or "関与" in purpose_text:
        return "経営関与"
    if "重要提案行為" in purpose_text:
        return "重要提案"
    return "その他"


def build_report_entry(doc, activists, xbrl_data=None, edinet_code_map=None):
    """EDINET ドキュメントから報告エントリを構築"""
    sec_code = extract_sec_code(doc.get("secCode", ""))
    filer_name = doc.get("filerName", "").strip()

    # アクティビスト判定
    matched_activist = match_activist(filer_name, activists)

    # 報告種別（docDescriptionから判定）
    doc_desc = doc.get("docDescription", "") or ""
    doc_type = doc.get("docTypeCode", "")
    if doc_type == DOC_TYPE_CORRECTION:
        report_type = "訂正報告"
    elif "変更報告" in doc_desc:
        report_type = "変更報告"
    elif "大量保有" in doc_desc:
        report_type = "新規報告"
    else:
        report_type = "その他"

    # 対象企業名の取得（優先順位: XBRL → EDINETコードリスト）
    target_name = ""

    # 1. XBRLから抽出した企業名（提出者名と同じ場合は除外）
    if xbrl_data and xbrl_data.get("target_company"):
        xbrl_name = xbrl_data["target_company"]
        # filer_name と一致 or 部分一致する場合はスキップ
        if xbrl_name not in filer_name and filer_name not in xbrl_name:
            target_name = xbrl_name

    # 2. EDINETコードリストから企業名
    if not target_name and edinet_code_map:
        issuer_code = (doc.get("issuerEdinetCode") or "").strip()
        subject_code = (doc.get("subjectEdinetCode") or "").strip()
        target_name = edinet_code_map.get(issuer_code, "") or edinet_code_map.get(subject_code, "")

    # 3. XBRLからの証券コードで補完
    if xbrl_data and xbrl_data.get("sec_code") and not sec_code:
        sec_code = extract_sec_code(xbrl_data["sec_code"])

    entry = {
        "doc_id": doc.get("docID", ""),
        "date": doc.get("submitDateTime", "")[:10],
        "filer_name": filer_name,
        "issuer_name": doc.get("issuerEdinetCode", ""),
        "sec_code": sec_code,
        "target_company": target_name if target_name else "",
        "report_type": report_type,
        "edinet_url": f"https://disclosure2.edinet-fsa.go.jp/WZEK0040.aspx?{doc.get('docID', '')}",
    }

    # XBRL から抽出したデータを追加
    if xbrl_data:
        if "holding_ratio" in xbrl_data:
            entry["holding_ratio"] = xbrl_data["holding_ratio"]
        if "purpose" in xbrl_data:
            entry["purpose"] = xbrl_data["purpose"]
        if "purpose_detail" in xbrl_data:
            entry["purpose_detail"] = xbrl_data["purpose_detail"]

    # アクティビスト / 注目投資家情報を追加
    if matched_activist:
        investor_type = matched_activist.get("type", "activist")
        if investor_type == "notable_holder":
            entry["is_activist"] = False
            entry["is_notable"] = True
        else:
            entry["is_activist"] = True
            entry["is_notable"] = False
        entry["activist_id"] = matched_activist["id"]
        entry["activist_type"] = investor_type
    else:
        entry["is_activist"] = False
        entry["is_notable"] = False

    return entry


def load_existing_data():
    """既存のデータファイルを読み込む"""
    if OUTPUT_FILE.exists():
        try:
            with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return {"reports": [], "last_updated": "", "activists_meta": {}}


def merge_reports(existing_reports, new_reports):
    """既存データと新規データをマージ（重複排除）"""
    seen_ids = set()
    merged = []

    # 新しいデータを優先
    for report in new_reports + existing_reports:
        doc_id = report.get("doc_id", "")
        if doc_id and doc_id not in seen_ids:
            seen_ids.add(doc_id)
            merged.append(report)

    # 日付の新しい順にソート
    merged.sort(key=lambda r: r.get("date", ""), reverse=True)
    return merged


def build_activist_summary(reports, activists):
    """アクティビスト・注目投資家別の保有銘柄サマリーを構築"""
    activist_holdings = {}

    for report in reports:
        if not report.get("is_activist") and not report.get("is_notable"):
            continue

        activist_id = report.get("activist_id", "")
        if not activist_id:
            continue

        if activist_id not in activist_holdings:
            # 元データから基本情報を取得
            base_info = next(
                (a for a in activists if a["id"] == activist_id), {}
            )
            activist_holdings[activist_id] = {
                "id": activist_id,
                "name": base_info.get("name", report.get("filer_name", "")),
                "type": base_info.get("type", "fund"),
                "representative": base_info.get("representative", ""),
                "headquarters": base_info.get("headquarters", ""),
                "description": base_info.get("description", ""),
                "focus_sectors": base_info.get("focus_sectors", []),
                "holdings": [],
                "report_count": 0,
                "latest_date": "",
            }

        entry = activist_holdings[activist_id]
        entry["report_count"] += 1

        if report["date"] > entry.get("latest_date", ""):
            entry["latest_date"] = report["date"]

        # 保有銘柄リストに追加（最新の報告のみ保持）
        sec_code = report.get("sec_code", "")
        existing = next(
            (h for h in entry["holdings"] if h.get("sec_code") == sec_code),
            None,
        )
        if existing:
            if report["date"] >= existing.get("date", ""):
                existing.update({
                    "date": report["date"],
                    "holding_ratio": report.get("holding_ratio"),
                    "purpose": report.get("purpose", ""),
                    "report_type": report.get("report_type", ""),
                })
        else:
            entry["holdings"].append({
                "sec_code": sec_code,
                "filer_name": report.get("filer_name", ""),
                "target_company": report.get("target_company", ""),
                "date": report["date"],
                "holding_ratio": report.get("holding_ratio"),
                "purpose": report.get("purpose", ""),
                "report_type": report.get("report_type", ""),
            })

    # report_count でソート
    return dict(
        sorted(
            activist_holdings.items(),
            key=lambda x: x[1]["report_count"],
            reverse=True,
        )
    )


def main():
    print("=" * 60)
    print(f"EDINET 大量保有報告書 取得スクリプト")
    print(f"実行時刻: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S JST')}")
    print("=" * 60)

    if not API_KEY:
        print("[ERROR] EDINET_API_KEY が設定されていません。", file=sys.stderr)
        print("  環境変数 EDINET_API_KEY を設定してください。", file=sys.stderr)
        print("  取得先: https://disclosure2dl.edinet-fsa.go.jp/", file=sys.stderr)
        # API キーなしでもサンプルデータで動作確認可能にする
        print("[INFO] サンプルデータで出力を生成します。")
        generate_sample_data()
        return

    # 既知アクティビストを読み込み
    activists = load_known_activists()
    print(f"既知アクティビスト: {len(activists)} 件")

    # EDINET コードリスト（企業名マッピング）
    edinet_code_map = load_edinet_code_map()

    # 既存データを読み込み
    existing_data = load_existing_data()
    existing_reports = existing_data.get("reports", [])
    print(f"既存データ: {len(existing_reports)} 件")

    # 対象日付を算出（既存データがあれば増分取得、なければフル取得）
    today = datetime.now(JST).date()
    scan_days = INCREMENTAL_DAYS if existing_reports else LOOKBACK_DAYS
    print(f"取得モード: {'増分' if existing_reports else '初回フル'}（{scan_days}日分）")
    dates = [(today - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(scan_days)]

    new_reports = []

    for i, date_str in enumerate(dates):
        print(f"\r  取得中: {date_str} ({i + 1}/{len(dates)})", end="", flush=True)

        documents = fetch_document_list(date_str)
        if not documents:
            time.sleep(0.5)
            continue

        holdings = filter_large_holdings(documents)

        for doc in holdings:
            doc_id = doc.get("docID", "")

            # 既存データにあればスキップ
            if any(r.get("doc_id") == doc_id for r in existing_reports):
                continue

            # アクティビスト・注目投資家判定で優先的にXBRLダウンロード
            filer_name = doc.get("filerName", "")
            matched = match_activist(filer_name, activists)

            xbrl_data = {}
            if matched and API_KEY:
                time.sleep(1)  # レート制限対策
                xbrl_data = download_xbrl_and_extract(doc_id)

            entry = build_report_entry(doc, activists, xbrl_data, edinet_code_map)
            new_reports.append(entry)

        time.sleep(0.5)  # レート制限対策

    print(f"\n新規取得: {new_reports.__len__()} 件")

    # 既存データの edinet_url 修復（S100 二重付与の修正）
    for r in existing_reports:
        url = r.get("edinet_url", "")
        if "S100S100" in url:
            r["edinet_url"] = url.replace("S100S100", "S100")

    # マージ
    all_reports = merge_reports(existing_reports, new_reports)

    # アクティビスト別サマリー
    activist_summary = build_activist_summary(all_reports, activists)

    # 出力
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    output = {
        "last_updated": datetime.now(JST).strftime("%Y-%m-%dT%H:%M:%S+09:00"),
        "total_reports": len(all_reports),
        "activist_reports": sum(1 for r in all_reports if r.get("is_activist")),
        "reports": all_reports,
        "activists": activist_summary,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n出力完了: {OUTPUT_FILE}")
    print(f"  総報告件数: {len(all_reports)}")
    print(f"  アクティビスト関連: {output['activist_reports']}")
    print(f"  追跡中アクティビスト: {len(activist_summary)}")


def generate_sample_data():
    """API キーがない場合のサンプルデータ生成"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    # サンプルは最初の起動確認用。実際のデータはEDINET APIから取得される。
    sample = {
        "last_updated": datetime.now(JST).strftime("%Y-%m-%dT%H:%M:%S+09:00"),
        "total_reports": 0,
        "activist_reports": 0,
        "reports": [],
        "activists": {},
        "_note": "EDINET_API_KEY を設定するとリアルデータが取得されます"
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(sample, f, ensure_ascii=False, indent=2)

    print(f"サンプルデータ出力完了: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
