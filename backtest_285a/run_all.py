# -*- coding: utf-8 -*-
"""ワンコマンド実行: データ取得 → バックテスト → レポート生成。

使い方:
  python3 run_all.py               # 実データ取得から全実行(要ネットワーク)
  python3 run_all.py --offline     # 既存キャッシュのみで実行
  python3 run_all.py --synthetic   # 合成データで動作検証(出力にSYNTHETIC注記)
"""
import sys

import config
import yahoo_fetch
import backtest
import report
from data_model import Market


def main(argv):
    synthetic = "--synthetic" in argv
    offline = "--offline" in argv or synthetic

    if synthetic:
        import synth
        print("[1/3] 合成データ生成(SYNTHETIC)")
        synth.generate()
    else:
        print("[1/3] Yahoo Chart API からデータ取得")
        failures = yahoo_fetch.fetch_all(offline=offline)
        critical = [f for f in failures if f[0] == config.SYM_MAIN]
        if critical:
            print(f"致命的: {config.SYM_MAIN} のデータがありません: {critical}", file=sys.stderr)
            return 1
        if failures:
            print(f"警告: 一部銘柄の取得に失敗(該当機能は縮退): {failures}", file=sys.stderr)

    print("[2/3] バックテスト実行(walk-forward + 感度分析)")
    mkt = Market()
    out = backtest.run_all(mkt)

    print("[3/3] レポート生成")
    outdir = config.OUTPUT + ("_sample" if synthetic else "")
    report.generate_all(out, outdir, synthetic=synthetic)
    print(f"完了: {outdir}/results.csv, daily_replay/, report.md")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
