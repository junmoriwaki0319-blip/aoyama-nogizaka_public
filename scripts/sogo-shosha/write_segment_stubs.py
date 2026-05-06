#!/usr/bin/env python3
"""
processed/segment-mix-<ticker>.json の stub を生成。
各社の実際の開示セグメント名を反映 (HTML 実装フェーズで値を埋める想定)。
データソース: 各社決算説明会資料 (2025/3 期)
"""

import json
from pathlib import Path

PROC = Path(__file__).resolve().parents[2] / "data" / "sogo-shosha" / "processed"

SEGMENT_DEFS = {
    "8058": {
        "name": "三菱商事",
        "segments": ["地球環境エネルギー", "マテリアルソリューション", "金属資源",
                     "社会インフラ", "モビリティ", "食品産業", "S.L.C.", "電力ソリューション"],
        "resource_segments": ["地球環境エネルギー", "金属資源"],
        "_note": "2024 年度組織再編後の現行 8 セグメント。地球環境エネルギーは旧『天然ガス』『石油・化学ソリューション』を統合"
    },
    "8031": {
        "name": "三井物産",
        "segments": ["金属資源", "エネルギー", "機械・インフラ", "化学品",
                     "鉄鋼製品", "生活産業", "次世代・機能推進"],
        "resource_segments": ["金属資源", "エネルギー"],
    },
    "8001": {
        "name": "伊藤忠商事",
        "segments": ["繊維", "機械", "金属", "エネルギー・化学品",
                     "食料", "住生活", "情報・金融", "第8 (THE 8th)"],
        "resource_segments": ["金属", "エネルギー・化学品"],
    },
    "8053": {
        "name": "住友商事",
        "segments": ["鉄鋼", "自動車", "輸送機・建機", "都市総合開発",
                     "メディア・デジタル", "ライフスタイル", "資源",
                     "化学品・エレクトロニクス・農業", "エネルギートランスフォーメーション"],
        "resource_segments": ["資源"],
        "_note": "2024 年度組織再編後の現行 9 セグメント。エネルギートランスフォーメーションが脱炭素軸として独立"
    },
    "8002": {
        "name": "丸紅",
        "segments": ["生活産業", "情報・物流", "金属", "エネルギー・金属",
                     "電力", "インフラプロジェクト", "航空・船舶", "金融・リース・不動産",
                     "食料・農業", "アグリ事業"],
        "resource_segments": ["金属", "エネルギー・金属"],
    },
    "2768": {
        "name": "双日",
        "segments": ["金属・資源", "化学", "自動車", "インフラ・ヘルスケア",
                     "リテール・コンシューマーサービス"],
        "resource_segments": ["金属・資源"],
        "_note": "「資源 / 非資源」分解は IR で直接提示されない。本配列は決算説明会資料準拠"
    },
    "8015": {
        "name": "豊田通商",
        "segments": ["金属", "グローバル部品 & ロジスティクス", "自動車",
                     "機械・エネルギー・プラントプロジェクト",
                     "化学品 & エレクトロニクス", "食料 & 消費財", "アフリカ"],
        "resource_segments": ["金属"],
        "_note": "自動車関連が約 50% で支配的。他商社と性格が異なる"
    },
}


def main():
    PROC.mkdir(parents=True, exist_ok=True)
    for ticker, info in SEGMENT_DEFS.items():
        stub = {
            "schema_version": "0.1",
            "ticker": ticker,
            "name": info["name"],
            "_status": "STUB - 値未取得。各社決算説明会資料 (FY24) からセグメント別営業利益を抽出予定",
            "_source_target": "決算説明会資料 (2025/3 期)",
            "fiscal_years": ["fy21", "fy22", "fy23", "fy24", "fy25"],
            "segments": [
                {
                    "name": seg,
                    "is_resource": seg in info.get("resource_segments", []),
                    "operating_income_jpy_oku": [None] * 5,
                }
                for seg in info["segments"]
            ],
        }
        if "_note" in info:
            stub["_note"] = info["_note"]
        out = PROC / f"segment-mix-{ticker}.json"
        out.write_text(json.dumps(stub, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  -> {out}")


if __name__ == "__main__":
    main()
