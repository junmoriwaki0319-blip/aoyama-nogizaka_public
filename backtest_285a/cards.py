# -*- coding: utf-8 -*-
"""カード生成: 8:55スナップショットから当日のカード(A/B/C/D)を組む。

型:
  A 押し目買い : 支持帯上端で指値買い
  B 戻り売り   : 抵抗帯下端で指値売り(減衰印字がアーム条件)
  C 深部受け   : バースト下げ→減衰→V字印字で成行買い
  D 突破       : 帯上限の「出来高定着」超えで翌バー成行買い
"""
from dataclasses import dataclass, field

import config


@dataclass
class Card:
    kind: str            # "A"|"B"|"C"|"D"
    side: str            # "long"|"short"
    entry_level: float | None   # 指値/トリガー座標 (C/Dは動的、Noneあり)
    stop: float | None
    target: float | None
    stop_width: float
    note: str = ""
    # 執行状態(engineが書き込む)
    status: str = "pending"
    fill_ts: object = None
    fill_price: float = None
    exit_ts: object = None
    exit_price: float = None
    exit_reason: str = ""
    pnl: float = None
    mae: float = None
    mfe: float = None


def adjust_stop_for_round(stop: float, side: str) -> float:
    """丸数字±60円内の損切りは禁止 → 丸数字±70円へずらす(不利方向へ)。"""
    r = round(stop / config.ROUND_STEP) * config.ROUND_STEP
    if abs(stop - r) <= config.ROUND_FORBID:
        return r - config.ROUND_SHIFT if side == "long" else r + config.ROUND_SHIFT
    return stop


def _ref_price(snap, use_adr: bool) -> float:
    """カード座標の基準価格。ADRアンカーON時は予測ギャップの一部を織り込む。"""
    ref = snap.prev_close
    if use_adr and snap.adr_gap_pred is not None:
        if abs(snap.adr_gap_pred) >= snap.prev_close * config.ADR_GAP_MIN:
            ref = snap.prev_close + snap.adr_gap_pred * config.ADR_GAP_COEF
    return ref


def _support_resistance(snap, ref: float):
    """しこり帯・前日高安・5日高安・窓端から支持/抵抗レベル群を作る。"""
    supports = []   # (level, label) 帯は上端を使う
    resists = []    # 帯は下端を使う
    for lo, hi, share in snap.vp_bands:
        if hi <= ref:
            supports.append((hi, f"しこり帯上端{lo:.0f}-{hi:.0f}"))
        elif lo >= ref:
            resists.append((lo, f"しこり帯下端{lo:.0f}-{hi:.0f}"))
    for lv, lab in [(snap.prev_low, "前日安値"), (snap.lo5, "5日安値")]:
        if lv and lv < ref:
            supports.append((lv, lab))
    for lv, lab in [(snap.prev_high, "前日高値"), (snap.hi5, "5日高値")]:
        if lv and lv > ref:
            resists.append((lv, lab))
    for lo, hi, gdate in snap.gaps:
        if hi <= ref:
            supports.append((hi, f"未埋め窓上端({gdate})"))
        elif lo >= ref:
            resists.append((lo, f"未埋め窓下端({gdate})"))
    for m, lab in [(snap.ma5, "MA5"), (snap.ma25, "MA25"), (snap.ma75, "MA75")]:
        if m and m < ref:
            supports.append((m, lab))
        elif m and m > ref:
            resists.append((m, lab))
    supports.sort(key=lambda x: -x[0])   # refに近い順
    resists.sort(key=lambda x: x[0])
    return supports, resists


def _long_target(entry: float, width: float, resists: list) -> float:
    t = entry + config.TARGET_RR * width
    near = [lv for lv, _ in resists if lv > entry + width * 0.8]
    if near:
        t = min(t, min(near) - config.TARGET_LEVEL_BUFFER)
    return t


def _short_target(entry: float, width: float, supports: list) -> float:
    t = entry - config.TARGET_RR * width
    near = [lv for lv, _ in supports if lv < entry - width * 0.8]
    if near:
        t = max(t, max(near) + config.TARGET_LEVEL_BUFFER)
    return t


def generate_cards(snap, stop_width: float, use_adr: bool) -> list:
    """スナップショットから当日カード一式を生成(座標は8:55に確定)。"""
    ref = _ref_price(snap, use_adr)
    supports, resists = _support_resistance(snap, ref)
    cards = []

    # A: 押し目買い — 最有力支持(帯上端優先=リスト先頭)で指値
    if supports:
        lv, lab = supports[0]
        if lv > snap.limit_low:
            stop = adjust_stop_for_round(lv - stop_width, "long")
            cards.append(Card("A", "long", lv, stop, _long_target(lv, stop_width, resists),
                              stop_width, note=f"支持={lab}"))

    # B: 戻り売り — 最有力抵抗の下端で指値(減衰印字がアーム条件)
    if resists:
        lv, lab = resists[0]
        if lv < snap.limit_high:
            stop = adjust_stop_for_round(lv + stop_width, "short")
            cards.append(Card("B", "short", lv, stop, _short_target(lv, stop_width, supports),
                              stop_width, note=f"抵抗={lab}(減衰条件付)"))

    # C: 深部受け — 座標は動的(バースト→減衰→V字)。engineが検出時にstop/targetを確定
    cards.append(Card("C", "long", None, None, None, stop_width, note="バースト→減衰→V字"))

    # D: 突破 — 帯上限(=最有力抵抗)の出来高定着超え
    if resists:
        lv, lab = resists[0]
        cards.append(Card("D", "long", lv, None, None, stop_width, note=f"突破座標={lab}"))

    return cards
