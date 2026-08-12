# -*- coding: utf-8 -*-
"""Yahooキャッシュの正規化とポイント・イン・タイム(PIT)アクセス。

ルックアヘッド禁止の実装方針:
- 全バーはJST datetimeを持つ。`bars_before(ts)` 系のみを判断ロジックに使う。
- 日足はその日の「確定後」にのみ利用可能(スナップショットは前日以前の日足のみ参照)。
"""
import json
import os
from datetime import datetime, timedelta, timezone
from collections import namedtuple

import config
import yahoo_fetch

JST = timezone(timedelta(hours=9))

Bar = namedtuple("Bar", "ts open high low close volume")


def _load_raw(symbol: str, interval: str) -> dict | None:
    path = yahoo_fetch.cache_path(symbol, interval)
    if not os.path.exists(path):
        return None
    with open(path) as f:
        return json.load(f)


def load_bars(symbol: str, interval: str) -> list:
    """キャッシュから正規化済みバー(JST・None除去・時刻昇順)を返す。無ければ[]。"""
    raw = _load_raw(symbol, interval)
    if raw is None:
        return []
    result = raw["chart"]["result"][0]
    ts_list = result.get("timestamp") or []
    q = result["indicators"]["quote"][0]
    bars = []
    for i, ts in enumerate(ts_list):
        o, h, l, c = q["open"][i], q["high"][i], q["low"][i], q["close"][i]
        v = q["volume"][i]
        if o is None or h is None or l is None or c is None:
            continue
        bars.append(Bar(datetime.fromtimestamp(ts, JST), o, h, l, c, v or 0))
    bars.sort(key=lambda b: b.ts)
    return bars


class Market:
    """全銘柄のバーを保持し、PITアクセスを提供する。"""

    def __init__(self):
        self._cache = {}

    def bars(self, symbol: str, interval: str) -> list:
        key = (symbol, interval)
        if key not in self._cache:
            self._cache[key] = load_bars(symbol, interval)
        return self._cache[key]

    # ---- PIT ヘルパ ---------------------------------------------------
    def daily_before(self, symbol: str, date) -> list:
        """dateより前の日付の日足のみ(スナップショット用)。"""
        return [b for b in self.bars(symbol, "1d") if b.ts.date() < date]

    def last_bar_at_or_before(self, symbol: str, interval: str, ts: datetime):
        """ts以前に「バーが閉じている」最新バー。バー開始時刻+interval ≤ ts で判定。"""
        span = {"5m": 300, "1h": 3600, "1d": 86400}[interval]
        cand = None
        for b in self.bars(symbol, interval):
            if b.ts + timedelta(seconds=span) <= ts:
                cand = b
            else:
                break
        return cand

    def session_bars_5m(self, symbol: str, date) -> list:
        """指定日の5分足(JST日付一致)。"""
        return [b for b in self.bars(symbol, "5m") if b.ts.date() == date]

    def sessions_5m(self, symbol: str) -> list:
        """5分足が存在する営業日リスト(昇順)。"""
        seen = []
        for b in self.bars(symbol, "5m"):
            d = b.ts.date()
            if not seen or seen[-1] != d:
                if d not in seen:
                    seen.append(d)
        return sorted(seen)

    def daily_close_map(self, symbol: str) -> dict:
        return {b.ts.date(): b.close for b in self.bars(symbol, "1d")}

    def close_auction_slippage(self, symbol: str, date) -> float | None:
        """日足公式終値 − 最終5分バー終値 = 引けオークション滑り(円)。

        5分足は引けオークション(15:25-15:30)を捕捉しないため別途記録する。
        """
        dmap = self.daily_close_map(symbol)
        if date not in dmap:
            return None
        s = self.session_bars_5m(symbol, date)
        if not s:
            return None
        return dmap[date] - s[-1].close
