# -*- coding: utf-8 -*-
"""場中執行エンジン: 当日の5分足のみでカードの発動・約定・出口を判定する。

執行規約(仕様):
- 指値は「バー安値<指値」(買い)/「バー高値>指値」(売り)で約定
- 寄りバー(9:00-9:05)は約定禁止
- スリッページ: 成行・逆指値の執行価格に1tick(10円)不利方向で適用。指値は指値で約定
- OCO同バー両到達は損切り優先(保守)
- 時間規制ON: 14:30以降新規禁止 / 15:20強制手仕舞い。OFF: 大引け最終バー終値で手仕舞い
- 同時保有1枚・1日2枚まで(同時1枚はモデリング選択として報告書に明記)
"""
from datetime import datetime, timedelta

import config
import cards as cards_mod
from data_model import JST


def _t(date, hm):
    return datetime(date.year, date.month, date.day, hm[0], hm[1], tzinfo=JST)


class DayResult:
    def __init__(self, date):
        self.date = date
        self.trades = []      # 約定したCard
        self.cards = []       # 生成された全Card
        self.log = []         # 執行ログ(daily_replay用)
        self.kr_blocked_from = None


def _decay_streak(vols: list, n_bars: int) -> int:
    """直近から遡って減衰印字(60分ピークの1/3以下)が連続している本数。"""
    streak = 0
    for i in range(len(vols) - 1, 0, -1):
        lb = vols[max(0, i - config.DECAY_LOOKBACK_BARS):i]
        if not lb:
            break
        if vols[i] <= max(lb) * config.DECAY_RATIO:
            streak += 1
        else:
            break
    return streak


def _is_burst_down(bars, i) -> bool:
    b = bars[i]
    if b.close >= b.open:
        return False
    prev = bars[max(0, i - config.BURST_VOL_AVG_BARS):i]
    if len(prev) < 5:
        return False
    avg = sum(x.volume for x in prev) / len(prev)
    return (avg > 0 and b.volume > config.BURST_VOL_MULT * avg) or (b.high - b.low > config.BURST_RANGE)


def _breakout_confirmed(bars, i, level) -> bool:
    b = bars[i]
    if b.close <= level:
        return False
    prev = bars[max(0, i - config.BURST_VOL_AVG_BARS):i]
    if len(prev) < 5:
        return False
    avg = sum(x.volume for x in prev) / len(prev)
    return avg > 0 and b.volume >= config.BREAKOUT_VOL_MULT * avg


def _close_position(pos, ts, price, reason, log):
    pos.exit_ts, pos.exit_price, pos.exit_reason = ts, price, reason
    sign = 1.0 if pos.side == "long" else -1.0
    pos.pnl = sign * (price - pos.fill_price) * config.LOT
    pos.status = "closed"
    log.append(f"{ts:%H:%M} EXIT {pos.kind} {reason} @{price:.0f} pnl={pos.pnl:+.0f}円")


def run_day(bars, day_cards, snap, date, decay_n: int, time_rule: bool,
            kr_block_ts=None) -> DayResult:
    """1営業日の5分足リプレイ。bars=当日の285A 5分足(時刻昇順)。"""
    res = DayResult(date)
    res.cards = day_cards
    res.kr_blocked_from = kr_block_ts
    if not bars:
        return res

    open_bar_ts = _t(date, config.OPEN_BAR)
    no_new_ts = _t(date, config.NO_NEW_AFTER)
    force_close_ts = _t(date, config.FORCE_CLOSE)

    pos = None
    entries = 0
    vols = []
    c_state = {"burst_i": None, "decayed": False}
    b_armed = False
    d_pending_entry = False   # 前バーで突破確認→当バー寄りで成行
    cardA = next((c for c in day_cards if c.kind == "A"), None)
    cardB = next((c for c in day_cards if c.kind == "B"), None)
    cardC = next((c for c in day_cards if c.kind == "C"), None)
    cardD = next((c for c in day_cards if c.kind == "D"), None)

    for i, bar in enumerate(bars):
        vols.append(bar.volume)
        is_open_bar = bar.ts == open_bar_ts
        can_new = (
            entries < config.MAX_CARDS_PER_DAY
            and pos is None
            and not is_open_bar
            and not (time_rule and bar.ts >= no_new_ts)
            and not (kr_block_ts is not None and bar.ts >= kr_block_ts)
        )

        # ---- 強制手仕舞い ------------------------------------------------
        if time_rule and pos is not None and bar.ts >= force_close_ts:
            _close_position(pos, bar.ts, bar.open - (config.TICK if pos.side == "long" else -config.TICK),
                            "time_force_close", res.log)
            res.trades.append(pos)
            pos = None
            continue

        # ---- 建玉の出口(OCO: 損切り優先) --------------------------------
        if pos is not None and not is_open_bar:
            if pos.side == "long":
                pos.mae = min(pos.mae, bar.low - pos.fill_price)
                pos.mfe = max(pos.mfe, bar.high - pos.fill_price)
                if bar.low <= pos.stop:
                    _close_position(pos, bar.ts, pos.stop - config.TICK, "stop", res.log)
                    res.trades.append(pos); pos = None
                elif bar.high >= pos.target:
                    _close_position(pos, bar.ts, pos.target, "target", res.log)
                    res.trades.append(pos); pos = None
            else:
                pos.mae = min(pos.mae, pos.fill_price - bar.high)
                pos.mfe = max(pos.mfe, pos.fill_price - bar.low)
                if bar.high >= pos.stop:
                    _close_position(pos, bar.ts, pos.stop + config.TICK, "stop", res.log)
                    res.trades.append(pos); pos = None
                elif bar.low <= pos.target:
                    _close_position(pos, bar.ts, pos.target, "target", res.log)
                    res.trades.append(pos); pos = None
        elif pos is not None and is_open_bar:
            # 寄りバーは約定禁止だがMAE/MFEは記録(日跨ぎ玉なしのため通常到達しない)
            if pos.side == "long":
                pos.mae = min(pos.mae, bar.low - pos.fill_price)
                pos.mfe = max(pos.mfe, bar.high - pos.fill_price)
            else:
                pos.mae = min(pos.mae, pos.fill_price - bar.high)
                pos.mfe = max(pos.mfe, pos.fill_price - bar.low)

        # ---- D: 前バー突破確認 → 当バー寄りで成行 -----------------------
        if pos is None and d_pending_entry and cardD and cardD.status == "pending" and can_new:
            entry = bar.open + config.TICK
            stop = cards_mod.adjust_stop_for_round(entry - cardD.stop_width, "long")
            supports, resists = cards_mod._support_resistance(snap, entry)
            cardD.stop, cardD.target = stop, cards_mod._long_target(entry, cardD.stop_width, resists)
            cardD.status, cardD.fill_ts, cardD.fill_price = "open", bar.ts, entry
            cardD.mae = cardD.mfe = 0.0
            pos = cardD; entries += 1
            res.log.append(f"{bar.ts:%H:%M} ENTRY D 成行買い @{entry:.0f} stop={stop:.0f} tgt={cardD.target:.0f}")
        d_pending_entry = False

        # ---- 新規約定判定 ------------------------------------------------
        if can_new and pos is None:
            # A: 押し目買い指値
            if cardA and cardA.status == "pending" and bar.low < cardA.entry_level:
                cardA.status, cardA.fill_ts, cardA.fill_price = "open", bar.ts, cardA.entry_level
                cardA.mae = cardA.mfe = 0.0
                pos = cardA; entries += 1
                res.log.append(f"{bar.ts:%H:%M} ENTRY A 指値買い @{cardA.fill_price:.0f} "
                               f"stop={cardA.stop:.0f} tgt={cardA.target:.0f}")
                # 同バーで損切りも到達なら保守的に損切り(OCO同バー損切り優先の拡張)
                if bar.low <= cardA.stop:
                    _close_position(pos, bar.ts, cardA.stop - config.TICK, "stop_same_bar", res.log)
                    res.trades.append(pos); pos = None
            # B: 減衰印字がアーム条件の戻り売り指値
            elif cardB and cardB.status == "pending" and b_armed and bar.high > cardB.entry_level:
                cardB.status, cardB.fill_ts, cardB.fill_price = "open", bar.ts, cardB.entry_level
                cardB.mae = cardB.mfe = 0.0
                pos = cardB; entries += 1
                res.log.append(f"{bar.ts:%H:%M} ENTRY B 戻り売り @{cardB.fill_price:.0f} "
                               f"stop={cardB.stop:.0f} tgt={cardB.target:.0f}")
                if bar.high >= cardB.stop:
                    _close_position(pos, bar.ts, cardB.stop + config.TICK, "stop_same_bar", res.log)
                    res.trades.append(pos); pos = None
            # C: バースト→減衰→V字
            elif cardC and cardC.status == "pending" and c_state["decayed"] and i > 0 \
                    and bar.close > bars[i - 1].high:
                entry = bar.close + config.TICK
                stop = cards_mod.adjust_stop_for_round(entry - cardC.stop_width, "long")
                supports, resists = cards_mod._support_resistance(snap, entry)
                cardC.entry_level, cardC.stop = entry, stop
                cardC.target = cards_mod._long_target(entry, cardC.stop_width, resists)
                cardC.status, cardC.fill_ts, cardC.fill_price = "open", bar.ts, entry
                cardC.mae = cardC.mfe = 0.0
                pos = cardC; entries += 1
                res.log.append(f"{bar.ts:%H:%M} ENTRY C V字受け @{entry:.0f} stop={stop:.0f} tgt={cardC.target:.0f}")

        # ---- 状態更新(次バー以降の判定材料; 当バーの確定情報のみ使用) ----
        streak = _decay_streak(vols, decay_n)
        if cardB and not b_armed and streak >= decay_n:
            b_armed = True
            res.log.append(f"{bar.ts:%H:%M} B条件: 減衰印字 {streak}本連続 → 戻り売りアーム")
        if cardC:
            if c_state["burst_i"] is None and _is_burst_down(bars, i):
                c_state["burst_i"] = i
                res.log.append(f"{bar.ts:%H:%M} C条件: 下方バースト検出 (vol={bar.volume:.0f})")
            elif c_state["burst_i"] is not None and not c_state["decayed"] and streak >= decay_n:
                c_state["decayed"] = True
                res.log.append(f"{bar.ts:%H:%M} C条件: バースト後減衰 → V字待ち")
        if cardD and cardD.status == "pending" and not d_pending_entry \
                and _breakout_confirmed(bars, i, cardD.entry_level):
            d_pending_entry = True
            res.log.append(f"{bar.ts:%H:%M} D条件: 帯上限{cardD.entry_level:.0f}を出来高定着超え")

    # ---- 大引け(時間規制OFF時 or 未決済残)。最終バー終値で手仕舞い ------
    if pos is not None:
        last = bars[-1]
        _close_position(pos, last.ts, last.close - (config.TICK if pos.side == "long" else -config.TICK),
                        "eod_close", res.log)
        res.trades.append(pos)

    return res


def korea_signal_ts(mkt, date, delay_min: int):
    """000660.KS が寄りから-2%に達した時刻(+遅延)。KST=JST(UTC+9)。"""
    bars = mkt.session_bars_5m("000660.KS", date)
    if not bars:
        return None
    o = bars[0].open
    if o <= 0:
        return None
    for b in bars:
        if (b.close - o) / o <= config.KR_DROP_TH:
            # バー確定時刻 = 開始+5分。そこに遅延を加算
            return b.ts + timedelta(minutes=5 + delay_min)
    return None
