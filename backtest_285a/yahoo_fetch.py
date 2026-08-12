# -*- coding: utf-8 -*-
"""Yahoo Chart API (query1.finance.yahoo.com/v8/finance/chart) からのデータ取得。

- User-Agent 必須
- 取得結果は data_raw/{symbol}_{interval}.json に生レスポンスのままキャッシュ
- ネットワーク遮断環境では既存キャッシュのみで動作(--offline)
"""
import json
import os
import sys
import time
import urllib.request
import urllib.error

import config


def _url(symbol: str, interval: str, range_: str) -> str:
    from urllib.parse import quote
    return (
        f"https://query1.finance.yahoo.com/v8/finance/chart/{quote(symbol)}"
        f"?interval={interval}&range={range_}&includePrePost=false&events=div%2Csplit"
    )


def cache_path(symbol: str, interval: str) -> str:
    safe = symbol.replace("=", "_").replace("^", "_").replace(".", "_")
    return os.path.join(config.DATA_RAW, f"{safe}_{interval}.json")


def fetch_one(symbol: str, interval: str, range_: str, retries: int = 3) -> dict:
    req = urllib.request.Request(
        _url(symbol, interval, range_), headers={"User-Agent": config.USER_AGENT}
    )
    last_err = None
    for i in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
            last_err = e
            time.sleep(2 ** i)
    raise RuntimeError(f"fetch failed for {symbol} {interval}: {last_err}")


def fetch_all(offline: bool = False) -> list:
    """FETCH_SPECS の全銘柄を取得しキャッシュ。戻り値=取得失敗リスト。"""
    os.makedirs(config.DATA_RAW, exist_ok=True)
    failures = []
    for symbol, interval, range_ in config.FETCH_SPECS:
        path = cache_path(symbol, interval)
        if offline:
            if not os.path.exists(path):
                failures.append((symbol, interval, "no cache (offline)"))
            continue
        try:
            data = fetch_one(symbol, interval, range_)
            result = (data.get("chart") or {}).get("result")
            if not result:
                raise RuntimeError(f"empty chart result: {data.get('chart', {}).get('error')}")
            with open(path, "w") as f:
                json.dump(data, f)
            print(f"  fetched {symbol} {interval} -> {path}")
            time.sleep(0.5)  # 行儀よく
        except Exception as e:  # noqa: BLE001 - 続行して失敗一覧を返す
            failures.append((symbol, interval, str(e)))
            print(f"  FAILED {symbol} {interval}: {e}", file=sys.stderr)
    return failures


if __name__ == "__main__":
    offline = "--offline" in sys.argv
    fails = fetch_all(offline=offline)
    if fails:
        print("failures:", fails, file=sys.stderr)
        sys.exit(1)
