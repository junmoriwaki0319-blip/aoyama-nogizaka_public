# -*- coding: utf-8 -*-
"""バックテスト・オーケストレータ。

- 全営業日を8:55スナップショット→カード生成→場中リプレイで回す
- walk-forward: 前半でパラメータ探索、後半でOOS検証
- 感度分析: 損切り幅×減衰N×時間規制、ADRアンカー、韓国データ遅延、イベント層別
"""
from datetime import date as date_cls, datetime, timedelta

import config
import levels
import cards as cards_mod
import engine
from data_model import Market


# ---------------------------------------------------------------- 層別ヘルパ
def _parse(d: str):
    return datetime.strptime(d, "%Y-%m-%d").date()


def regime_of(d) -> str | None:
    for key, start, end, _label in config.REGIMES:
        if _parse(start) <= d <= _parse(end):
            return key
    return None


def event_class_of(d, sessions) -> str:
    """'event' / 'event+1' / 'normal'"""
    ev_dates = {_parse(x[0]) for x in config.EVENTS}
    if d in ev_dates:
        return "event"
    idx = sessions.index(d) if d in sessions else -1
    if idx > 0 and sessions[idx - 1] in ev_dates:
        return "event+1"
    # イベントが休日(米指標の現地日付等)の場合: 直後の営業日をevent+1扱い
    for ev in ev_dates:
        if ev not in sessions and idx >= 0:
            prevs = [s for s in sessions if s > ev]
            if prevs and prevs[0] == d:
                return "event+1"
    return "normal"


# ---------------------------------------------------------------- 1条件の実行
def run_variant(mkt: Market, sessions, stop_width, decay_n, time_rule,
                use_adr=True, kr_mode="realtime"):
    """条件一式で全営業日を回し DayResult のリストを返す。

    kr_mode: 'realtime' | 'delayed' | 'off'
    """
    results = []
    for d in sessions:
        snap = levels.build_snapshot(mkt, d, use_adr=use_adr)
        if snap is None:
            continue
        day_cards = cards_mod.generate_cards(snap, stop_width, use_adr)
        bars = mkt.session_bars_5m(config.SYM_MAIN, d)
        kr_ts = None
        if kr_mode == "realtime":
            kr_ts = engine.korea_signal_ts(mkt, d, 0)
        elif kr_mode == "delayed":
            kr_ts = engine.korea_signal_ts(mkt, d, config.KR_DELAY_MIN)
        day = engine.run_day(bars, day_cards, snap, d, decay_n, time_rule, kr_block_ts=kr_ts)
        day.snapshot = snap
        day.auction_slip = mkt.close_auction_slippage(config.SYM_MAIN, d)
        results.append(day)
    return results


# ---------------------------------------------------------------- 集計
def stats(day_results, day_filter=None, kind=None) -> dict:
    """発動率・勝率・平均RR・期待値。day_filter(date)->bool / kind でカード型層別。"""
    days = [r for r in day_results if day_filter is None or day_filter(r.date)]
    n_days = len(days)
    n_cards = sum(1 for r in days for c in r.cards if kind is None or c.kind == kind)
    trades = [t for r in days for t in r.trades
              if t.pnl is not None and (kind is None or t.kind == kind)]
    n = len(trades)
    wins = [t for t in trades if t.pnl > 0]
    losses = [t for t in trades if t.pnl <= 0]
    avg_win = sum(t.pnl for t in wins) / len(wins) if wins else 0.0
    avg_loss = sum(t.pnl for t in losses) / len(losses) if losses else 0.0
    return {
        "n_days": n_days,
        "n_cards": n_cards,
        "n_trades": n,
        "trigger_rate": n / n_cards if n_cards else 0.0,
        "win_rate": len(wins) / n if n else 0.0,
        "avg_win": avg_win,
        "avg_loss": avg_loss,
        "rr": (avg_win / abs(avg_loss)) if avg_loss else 0.0,
        "expectancy": sum(t.pnl for t in trades) / n if n else 0.0,
        "total_pnl": sum(t.pnl for t in trades),
        "reference_only": n < config.MIN_TRADES_FOR_CONCLUSION,
    }


# ---------------------------------------------------------------- 全体実行
def run_all(mkt: Market) -> dict:
    """メイン: walk-forward探索 + 感度分析一式。結果辞書を返す。"""
    sessions = mkt.sessions_5m(config.SYM_MAIN)
    if len(sessions) < 8:
        raise RuntimeError(f"5分足の営業日が不足: {len(sessions)}日")
    half = len(sessions) // 2
    is_days, oos_days = sessions[:half], sessions[half:]

    out = {"sessions": sessions, "is_days": is_days, "oos_days": oos_days}

    # ---- (1) パラメータ・グリッド(IS区間) -----------------------------
    grid = []
    for sw in config.STOP_WIDTH_GRID:
        for n in config.DECAY_N_GRID:
            for tr in (True, False):
                r = run_variant(mkt, is_days, sw, n, tr)
                s = stats(r)
                grid.append({"stop": sw, "decay_n": n, "time_rule": tr, "is": s})
    # 選択基準: n>=8なら期待値、未満なら総損益(その旨reportで明示)
    def _score(g):
        s = g["is"]
        return (s["expectancy"] if s["n_trades"] >= 8 else s["total_pnl"] / 10000.0)
    best = max(grid, key=_score)
    out["grid"] = grid
    out["best_params"] = {k: best[k] for k in ("stop", "decay_n", "time_rule")}

    # ---- (2) OOS検証 + 全期間(基準パラメータ) -------------------------
    bp = out["best_params"]
    oos_res = run_variant(mkt, oos_days, bp["stop"], bp["decay_n"], bp["time_rule"])
    out["oos_stats"] = stats(oos_res)
    full_res = run_variant(mkt, sessions, bp["stop"], bp["decay_n"], bp["time_rule"])
    out["full_results"] = full_res

    # ---- (2b) カード型別 -------------------------------------------------
    out["card_stats"] = {k: stats(full_res, kind=k) for k in ("A", "B", "C", "D")}

    # ---- (3) レジーム別(全期間混合は出さない) --------------------------
    out["regime_stats"] = {}
    for key, start, end, label in config.REGIMES:
        s, e = _parse(start), _parse(end)
        out["regime_stats"][key] = {
            "label": label, "range": (start, end),
            "stats": stats(full_res, lambda d, s=s, e=e: s <= d <= e),
        }

    # ---- (4) イベント層別 ------------------------------------------------
    out["event_stats"] = {}
    for cls in ("event", "event+1", "normal"):
        out["event_stats"][cls] = stats(
            full_res, lambda d, c=cls: event_class_of(d, sessions) == c)

    # ---- (5) 情報タイミング感度 -----------------------------------------
    # (a) ADRアンカー あり/なし
    adr_on = stats(full_res)
    adr_off_res = run_variant(mkt, sessions, bp["stop"], bp["decay_n"], bp["time_rule"],
                              use_adr=False)
    out["timing_adr"] = {"on": adr_on, "off": stats(adr_off_res)}

    # (b) 韓国データ: リアルタイム / 20分遅延 / 不使用
    out["timing_kr"] = {}
    for mode in ("realtime", "delayed", "off"):
        r = run_variant(mkt, sessions, bp["stop"], bp["decay_n"], bp["time_rule"],
                        kr_mode=mode)
        out["timing_kr"][mode] = stats(r)

    # ---- (6) 引けオークション滑り ---------------------------------------
    slips = [r.auction_slip for r in full_res if r.auction_slip is not None]
    out["auction_slip"] = {
        "n": len(slips),
        "mean": sum(slips) / len(slips) if slips else None,
        "abs_mean": sum(abs(x) for x in slips) / len(slips) if slips else None,
    }

    return out
