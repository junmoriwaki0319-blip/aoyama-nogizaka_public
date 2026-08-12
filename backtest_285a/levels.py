# -*- coding: utf-8 -*-
"""座標エンジン: 毎朝8:55スナップショット(前日以前の構造のみで再計算)。

除外情報(過去再現不能のため): 引け残り板・PTS・群衆情報。report.mdに明記される。
"""
from datetime import datetime, timedelta
from dataclasses import dataclass, field

import config
import jpx_limits
from data_model import JST, Market


@dataclass
class Snapshot:
    date: object                      # 対象営業日 (datetime.date)
    prev_high: float = None
    prev_low: float = None
    prev_close: float = None
    hi5: float = None                 # 直近5日高値
    lo5: float = None                 # 直近5日安値
    gaps: list = field(default_factory=list)      # 未埋め窓 [(下端, 上端, 発生日)]
    vp_bands: list = field(default_factory=list)  # しこり帯 [(帯下端, 帯上端, 出来高シェア)]
    ma5: float = None
    ma25: float = None
    ma75: float = None
    round_levels: list = field(default_factory=list)
    adr_theoretical: float = None     # KXIAY前日終値×USDJPY(8:55)×10
    adr_gap_pred: float = None        # 予測ギャップ(円) = 理論値 − 前日終値
    usdjpy_855: float = None
    niy_855: float = None
    limit_low: float = None           # 値幅制限(ストップ安)
    limit_high: float = None          # 値幅制限(ストップ高)
    notes: list = field(default_factory=list)


def _unfilled_gaps(daily: list) -> list:
    """日足列から未埋め窓を検出。

    窓: 当日安値 > 前日高値 (ギャップアップ) / 当日高値 < 前日安値 (ギャップダウン)。
    以後の日足レンジが窓域に触れた分だけ埋まったとみなし、未埋め部分のみ返す。
    """
    gaps = []
    for i in range(1, len(daily)):
        prev, cur = daily[i - 1], daily[i]
        if cur.low > prev.high:
            gaps.append([prev.high, cur.low, cur.ts.date()])   # up-gap [下端, 上端]
        elif cur.high < prev.low:
            gaps.append([cur.high, prev.low, cur.ts.date()])   # down-gap
        # 既存窓の埋まり判定
        for g in gaps:
            lo, hi = g[0], g[1]
            if lo >= hi:
                continue
            if cur.ts.date() <= g[2]:
                continue
            # 当日レンジと窓域の重なりを消し込む(上から/下から侵食)
            if cur.low <= lo and cur.high >= hi:
                g[0], g[1] = 0.0, 0.0
            elif lo <= cur.high <= hi:
                g[0] = cur.high
            elif lo <= cur.low <= hi:
                g[1] = cur.low
    return [(g[0], g[1], g[2]) for g in gaps if g[1] - g[0] > 1.0]


def _volume_profile(bars5m: list) -> list:
    """5分足VWAP近似(典型価格)×出来高の価格帯別集計 → 上位帯。"""
    if not bars5m:
        return []
    hist = {}
    total = 0.0
    for b in bars5m:
        tp = (b.high + b.low + b.close) / 3.0
        k = int(tp // config.VP_BIN)
        hist[k] = hist.get(k, 0.0) + b.volume
        total += b.volume
    if total <= 0:
        return []
    top = sorted(hist.items(), key=lambda kv: -kv[1])[: config.VP_TOP_BANDS * 2]
    # 隣接ビンをマージして帯にする
    top_keys = sorted(k for k, _ in top[: config.VP_TOP_BANDS * 2])
    bands = []
    start = prev = None
    vol = 0.0
    for k in top_keys:
        if start is None:
            start, prev, vol = k, k, hist[k]
        elif k == prev + 1:
            prev = k
            vol += hist[k]
        else:
            bands.append((start * config.VP_BIN, (prev + 1) * config.VP_BIN, vol / total))
            start, prev, vol = k, k, hist[k]
    if start is not None:
        bands.append((start * config.VP_BIN, (prev + 1) * config.VP_BIN, vol / total))
    bands.sort(key=lambda b: -b[2])
    return bands[: config.VP_TOP_BANDS]


def build_snapshot(mkt: Market, date, use_adr: bool = True) -> Snapshot | None:
    """dateの朝8:55時点スナップショット。前日以前のデータのみ使用。"""
    snap = Snapshot(date=date)
    daily = mkt.daily_before(config.SYM_MAIN, date)
    if len(daily) < 2:
        return None
    prev = daily[-1]
    snap.prev_high, snap.prev_low, snap.prev_close = prev.high, prev.low, prev.close
    last5 = daily[-5:]
    snap.hi5 = max(b.high for b in last5)
    snap.lo5 = min(b.low for b in last5)
    snap.gaps = _unfilled_gaps(daily[-120:])

    closes = [b.close for b in daily]
    for n, attr in [(5, "ma5"), (25, "ma25"), (75, "ma75")]:
        if len(closes) >= n:
            setattr(snap, attr, sum(closes[-n:]) / n)

    # しこり帯: 過去20営業日の5分足(全て date より前のセッション)
    sessions = [d for d in mkt.sessions_5m(config.SYM_MAIN) if d < date][-config.VP_DAYS:]
    bars5 = [b for b in mkt.bars(config.SYM_MAIN, "5m") if b.ts.date() in set(sessions)]
    snap.vp_bands = _volume_profile(bars5)

    # 丸数字(前日終値±2000円圏)
    base = int(snap.prev_close // config.ROUND_STEP)
    snap.round_levels = [
        (base + k) * config.ROUND_STEP for k in (-2, -1, 0, 1, 2)
        if (base + k) * config.ROUND_STEP > 0
    ]

    # 値幅制限(基準値段=前日終値)
    snap.limit_low, snap.limit_high = jpx_limits.daily_limits(snap.prev_close)

    # ADR理論値(寄りアンカー)
    ts855 = datetime(date.year, date.month, date.day, *config.SNAPSHOT_TIME, tzinfo=JST)
    if use_adr:
        adr_daily = mkt.daily_before(config.SYM_ADR, date)
        fx = mkt.last_bar_at_or_before(config.SYM_FX, "1h", ts855)
        if adr_daily and fx:
            snap.usdjpy_855 = fx.close
            snap.adr_theoretical = adr_daily[-1].close * fx.close * config.ADR_MULT
            snap.adr_gap_pred = snap.adr_theoretical - snap.prev_close

    niy = mkt.last_bar_at_or_before(config.SYM_NIY, "5m", ts855)
    if niy:
        snap.niy_855 = niy.close

    snap.notes.append("除外情報: 引け残り板・PTS・群衆情報(過去再現不能のため不使用)")
    return snap
