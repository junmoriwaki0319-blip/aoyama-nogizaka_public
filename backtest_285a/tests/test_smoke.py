# -*- coding: utf-8 -*-
"""スモークテスト: 合成データでパイプライン全体+個別ルールの検証。

実行: cd backtest_285a && python3 -m unittest discover tests
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config
import cards
import jpx_limits
from engine import _decay_streak


class TestRules(unittest.TestCase):
    def test_jpx_limit_width(self):
        # JPX制限値幅表の代表値
        self.assertEqual(jpx_limits.price_limit_width(2500), 500)
        self.assertEqual(jpx_limits.price_limit_width(4000), 700)
        self.assertEqual(jpx_limits.price_limit_width(8000), 1500)
        lo, hi = jpx_limits.daily_limits(3000)
        self.assertEqual((lo, hi), (2300, 3700))  # 基準3000は「3000以上5000未満」区分=700円

    def test_round_number_stop_forbidden_zone(self):
        # 丸数字±60円内は禁止 → ±70円へずらす
        self.assertEqual(cards.adjust_stop_for_round(4050.0, "long"), 3930.0)
        self.assertEqual(cards.adjust_stop_for_round(3950.0, "long"), 3930.0)
        self.assertEqual(cards.adjust_stop_for_round(4050.0, "short"), 4070.0)
        # 禁止帯の外はそのまま
        self.assertEqual(cards.adjust_stop_for_round(4200.0, "long"), 4200.0)

    def test_decay_streak(self):
        # 直近60分ピーク100万株 → 1/3以下が2本連続
        vols = [1_000_000] * 10 + [300_000, 250_000]
        self.assertGreaterEqual(_decay_streak(vols, 2), 2)
        vols2 = [1_000_000] * 10 + [900_000]
        self.assertEqual(_decay_streak(vols2, 2), 0)


class TestPipeline(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import synth
        synth.generate(seed=42)

    def test_full_pipeline(self):
        import backtest
        import report
        from data_model import Market
        mkt = Market()
        out = backtest.run_all(mkt)
        # walk-forward分割とレポート必須キー
        self.assertGreater(len(out["is_days"]), 0)
        self.assertGreater(len(out["oos_days"]), 0)
        self.assertLess(out["is_days"][-1], out["oos_days"][0])
        self.assertEqual(len(out["grid"]), len(config.STOP_WIDTH_GRID) * len(config.DECAY_N_GRID) * 2)
        # 出力ファイル
        outdir = os.path.join(config.BASE_DIR, "output_test")
        report.generate_all(out, outdir, synthetic=True)
        self.assertTrue(os.path.exists(os.path.join(outdir, "results.csv")))
        self.assertTrue(os.path.exists(os.path.join(outdir, "report.md")))
        self.assertTrue(os.listdir(os.path.join(outdir, "daily_replay")))

    def test_point_in_time_snapshot(self):
        """スナップショットが当日データを一切使わないこと(PIT検証)。"""
        import levels
        from data_model import Market
        mkt = Market()
        sessions = mkt.sessions_5m(config.SYM_MAIN)
        d = sessions[-1]
        snap = levels.build_snapshot(mkt, d)
        daily = mkt.bars(config.SYM_MAIN, "1d")
        todays = [b for b in daily if b.ts.date() == d]
        # 前日終値は当日の日足と無関係(前日以前の最後の足と一致)
        prevs = [b for b in daily if b.ts.date() < d]
        self.assertEqual(snap.prev_close, prevs[-1].close)
        if todays:
            self.assertNotEqual(snap.prev_close, todays[0].close)


if __name__ == "__main__":
    unittest.main()
