# -*- coding: utf-8 -*-
"""合成データ生成器(パイプライン検証専用)。

⚠️ ここで生成されるデータは全て人工の乱数系列であり、実相場ではない。
用途: ネットワーク遮断環境でのスモークテスト、および出力フォーマットのサンプル生成。
生成物には必ず SYNTHETIC の注記が付く(report.py側で制御)。

レジーム構造は依頼仕様を模倣: 7/17-8/1に急落局面(韓国連動)、8/3以降レンジ回復。
乱数は固定シードで決定的。
"""
import json
import os
import random
from datetime import datetime, timedelta, timezone

import config
import yahoo_fetch

JST = timezone(timedelta(hours=9))


def _sessions(start, end):
    d = start
    out = []
    while d <= end:
        if d.weekday() < 5:
            out.append(d)
        d += timedelta(days=1)
    return out


def _to_yahoo_json(ts, o, h, l, c, v):
    return {"chart": {"result": [{
        "meta": {"symbol": "SYNTH"},
        "timestamp": ts,
        "indicators": {"quote": [{"open": o, "high": h, "low": l, "close": c, "volume": v}]},
    }], "error": None}}


def _write(symbol, interval, payload):
    os.makedirs(config.DATA_RAW, exist_ok=True)
    with open(yahoo_fetch.cache_path(symbol, interval), "w") as f:
        json.dump(payload, f)


def _drift_for(d):
    """レジーム: 7/17-8/1 急落, 8/3〜 レンジ。"""
    if datetime(2026, 7, 17).date() <= d <= datetime(2026, 8, 1).date():
        return -0.012
    return 0.0005


def generate(seed=42):
    rng = random.Random(seed)
    end = datetime(2026, 8, 12).date()
    sessions_5m = _sessions(end - timedelta(days=85), end)[-60:]
    sessions_daily = _sessions(datetime(2024, 12, 18).date(), end)

    # ---- 285A 日足(上場来) + 5分足(60日) --------------------------------
    px = 1440.0
    ts_d, o_d, h_d, l_d, c_d, v_d = [], [], [], [], [], []
    daily_px = {}
    for d in sessions_daily:
        drift = _drift_for(d)
        # 上場来の長期トレンドをざっくり付与(検証用: 2026年半ばに8000円前後へ。
        # 損切り幅250-450円・バースト値幅300円という仕様と整合する価格帯にする)
        base_drift = 0.0046 if d < datetime(2026, 6, 1).date() else drift
        o = px * (1 + rng.gauss(base_drift * 0.3, 0.01))
        c = o * (1 + rng.gauss(base_drift, 0.022))
        h = max(o, c) * (1 + abs(rng.gauss(0, 0.008)))
        l = min(o, c) * (1 - abs(rng.gauss(0, 0.008)))
        v = int(rng.uniform(5e6, 2e7))
        t = int(datetime(d.year, d.month, d.day, 9, 0, tzinfo=JST).timestamp())
        ts_d.append(t); o_d.append(round(o)); h_d.append(round(h))
        l_d.append(round(l)); c_d.append(round(c)); v_d.append(v)
        daily_px[d] = (round(o), round(c))
        px = c
    _write(config.SYM_MAIN, "1d", _to_yahoo_json(ts_d, o_d, h_d, l_d, c_d, v_d))

    ts5, o5, h5, l5, c5, v5 = [], [], [], [], [], []
    for d in sessions_5m:
        if d not in daily_px:
            continue
        day_o, day_c = daily_px[d]
        p = float(day_o)
        n_bars = 78  # 9:00〜15:25発の5分足(昼休みは下でスキップ)
        for k in range(n_bars):
            t = datetime(d.year, d.month, d.day, 9, 0, tzinfo=JST) + timedelta(minutes=5 * k)
            if t.hour == 11 and t.minute >= 30 or t.hour == 12 and t.minute < 30:
                continue
            target_pull = (day_c - p) / max(1, n_bars - k) * 0.6
            o = p
            c = p + target_pull + rng.gauss(0, p * 0.006)
            h = max(o, c) + abs(rng.gauss(0, p * 0.003))
            l = min(o, c) - abs(rng.gauss(0, p * 0.003))
            v = int(abs(rng.gauss(2.2e5, 1.4e5))) + 1000
            if rng.random() < 0.03:
                v *= 4  # 時折バースト
                l -= abs(rng.gauss(0, p * 0.012))
            ts5.append(int(t.timestamp())); o5.append(round(o)); h5.append(round(h))
            l5.append(round(l)); c5.append(round(c)); v5.append(v)
            p = c
    _write(config.SYM_MAIN, "5m", _to_yahoo_json(ts5, o5, h5, l5, c5, v5))

    # ---- KXIAY 日足(285Aに概ね連動・ドル建て) ---------------------------
    ts_a, o_a, h_a, l_a, c_a, v_a = [], [], [], [], [], []
    for d in sessions_daily[-260:]:
        if d not in daily_px:
            continue
        c_jpy = daily_px[d][1]
        adr = c_jpy / 148.0 / config.ADR_MULT * (1 + rng.gauss(0.001, 0.012))
        # 米国市場の引け(≈JST翌朝5時)を前営業日の日付キーで記録
        t = int((datetime(d.year, d.month, d.day, 23, 0, tzinfo=JST)).timestamp())
        ts_a.append(t); o_a.append(adr); h_a.append(adr * 1.01)
        l_a.append(adr * 0.99); c_a.append(adr); v_a.append(int(rng.uniform(1e5, 5e5)))
    _write(config.SYM_ADR, "1d", _to_yahoo_json(ts_a, o_a, h_a, l_a, c_a, v_a))
    for sym in ("SNDK", "MU"):
        _write(sym, "1d", _to_yahoo_json(ts_a, o_a, h_a, l_a, c_a, v_a))
        _write(sym, "5m", _to_yahoo_json([], [], [], [], [], []))
    _write(config.SYM_ADR, "5m", _to_yahoo_json([], [], [], [], [], []))

    # ---- USDJPY 1h ------------------------------------------------------
    ts_f, c_f = [], []
    fx = 147.0
    t0 = datetime(sessions_daily[-90].year, sessions_daily[-90].month,
                  sessions_daily[-90].day, 0, 0, tzinfo=JST)
    for k in range(90 * 24):
        fx *= (1 + rng.gauss(0, 0.0008))
        ts_f.append(int((t0 + timedelta(hours=k)).timestamp())); c_f.append(fx)
    _write(config.SYM_FX, "1h", _to_yahoo_json(ts_f, c_f, c_f, c_f, c_f, [0] * len(ts_f)))

    # ---- 韓国(000660.KS, ^KS11) 5分+日足: 急落期に-3%日を混ぜる ----------
    for sym in config.SYM_KR:
        tsk, ok, hk, lk, ck, vk = [], [], [], [], [], []
        base = 200000.0 if sym == "000660.KS" else 2700.0
        for d in sessions_5m:
            crash = datetime(2026, 7, 17).date() <= d <= datetime(2026, 8, 1).date()
            day_ret = rng.gauss(-0.02 if crash and rng.random() < 0.5 else 0.0, 0.012)
            p = base
            for k in range(60):
                t = datetime(d.year, d.month, d.day, 9, 0, tzinfo=JST) + timedelta(minutes=5 * k)
                c = p * (1 + day_ret / 60 + rng.gauss(0, 0.001))
                tsk.append(int(t.timestamp())); ok.append(p); hk.append(max(p, c) * 1.001)
                lk.append(min(p, c) * 0.999); ck.append(c); vk.append(int(rng.uniform(1e4, 1e5)))
                p = c
            base = p
        _write(sym, "5m", _to_yahoo_json(tsk, ok, hk, lk, ck, vk))
        _write(sym, "1d", _to_yahoo_json([], [], [], [], [], []))

    # ---- NIY=F 5分(夜間も動く簡略版: 8:00-9:00のみ生成) ------------------
    tsn, cn = [], []
    for d in sessions_5m:
        p = daily_px.get(d, (40000, 40000))[0] * 13.0
        for k in range(12):
            t = datetime(d.year, d.month, d.day, 8, 0, tzinfo=JST) + timedelta(minutes=5 * k)
            p *= (1 + rng.gauss(0, 0.0005))
            tsn.append(int(t.timestamp())); cn.append(p)
    _write(config.SYM_NIY, "5m", _to_yahoo_json(tsn, cn, cn, cn, cn, [0] * len(tsn)))

    print(f"synthetic data written to {config.DATA_RAW} (seed={seed}, "
          f"5m sessions={len(sessions_5m)})")


if __name__ == "__main__":
    generate()
