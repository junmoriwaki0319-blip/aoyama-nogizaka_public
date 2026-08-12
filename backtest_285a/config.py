# -*- coding: utf-8 -*-
"""285A.T デイトレカード・バックテスト 設定。

全てのパラメータ・モデリング上の選択はこのファイルに集約する。
ヒューリスティックな定数(帯のビン幅、韓国リスクオフ閾値など)は
report.md にも「モデリング選択」として明記される。
"""

# ---------------------------------------------------------------- シンボル
SYM_MAIN = "285A.T"          # キオクシアHD
SYM_ADR = "KXIAY"            # キオクシアADR (1ADR = 0.1株 → 理論値 = ADR×USDJPY×10)
SYM_PEERS_US = ["SNDK", "MU"]
SYM_FX = "USDJPY=X"
SYM_KR = ["000660.KS", "^KS11"]   # SKハイニックス / KOSPI
SYM_NIY = "NIY=F"            # CME日経円建て先物

# Yahoo v8 chart API 取得仕様: (symbol, interval, range)
FETCH_SPECS = [
    (SYM_MAIN, "5m", "60d"),
    (SYM_MAIN, "1d", "max"),
    (SYM_ADR, "1d", "1y"),
    (SYM_ADR, "5m", "60d"),
    ("SNDK", "1d", "1y"),
    ("SNDK", "5m", "60d"),
    ("MU", "1d", "1y"),
    ("MU", "5m", "60d"),
    (SYM_FX, "1h", "180d"),
    ("000660.KS", "1d", "1y"),
    ("000660.KS", "5m", "60d"),
    ("^KS11", "1d", "1y"),
    ("^KS11", "5m", "60d"),
    (SYM_NIY, "5m", "60d"),
]

USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

# ---------------------------------------------------------------- 時刻 (JST)
SNAPSHOT_TIME = (8, 55)      # 毎朝のスナップショット時刻
OPEN_BAR = (9, 0)            # 寄りバー(9:00-9:05)は約定禁止
NO_NEW_AFTER = (14, 30)      # 時間規制ON時: 以降新規禁止
FORCE_CLOSE = (15, 20)       # 時間規制ON時: 強制手仕舞い
SESSION_END = (15, 30)       # 現物立会終了(2024/11〜 15:30)

# ---------------------------------------------------------------- カード規格
LOT = 100                    # サイズ100株固定
MAX_CARDS_PER_DAY = 2        # 1日2枚まで
TICK = 10.0                  # スリッページ 1tick = 10円(仕様固定)

STOP_WIDTH_BASE = 350.0      # 損切り幅の基準
STOP_WIDTH_GRID = [250.0, 350.0, 450.0]   # 感度分析
ROUND_STEP = 1000.0          # 丸数字 = 1,000円刻み
ROUND_FORBID = 60.0          # 丸数字±60円内に損切りを置くの禁止
ROUND_SHIFT = 70.0           # 禁止帯に掛かった場合のずらし先(丸数字±70円)

DECAY_N_BASE = 2             # 減衰印字: N本連続(基準2、感度1〜3)
DECAY_N_GRID = [1, 2, 3]
DECAY_RATIO = 1.0 / 3.0      # 5分出来高が直近60分ピークの1/3以下
DECAY_LOOKBACK_BARS = 12     # 直近60分 = 5分足12本

BURST_VOL_MULT = 3.0         # バースト: 出来高が20本平均の3倍超
BURST_VOL_AVG_BARS = 20
BURST_RANGE = 300.0          # または5分値幅300円超

BREAKOUT_VOL_MULT = 1.5     # D突破: 「出来高定着」= 終値が帯上限超え かつ 出来高≥20本平均×1.5

TARGET_RR = 2.0              # 利確: 損切り幅×2.0(次抵抗が近ければそちら優先)
TARGET_LEVEL_BUFFER = 10.0   # 次抵抗/支持レベル手前のバッファ

# しこり帯(価格帯別出来高)
VP_DAYS = 20                 # 過去20営業日
VP_BIN = 25.0                # ビン幅25円
VP_TOP_BANDS = 3             # 上位帯の数

# ADRアンカー: 予測ギャップの反映係数(感度(a)でON/OFF)
ADR_MULT = 10.0              # 理論値 = KXIAY前日終値×USDJPY×10
ADR_GAP_COEF = 0.5           # カード座標を予測ギャップ×係数だけシフト
ADR_GAP_MIN = 0.01           # 予測ギャップ1%未満は無視

# 韓国リスクオフ・フィルター(感度(b))
KR_DROP_TH = -0.02           # 000660.KS が寄りから-2%で当日新規停止
KR_DELAY_MIN = 20            # 遅延モデル: シグナル到達を20分遅らせる

# ---------------------------------------------------------------- 期間・レジーム
# レジーム分割(全期間混合の平均は出さない)
REGIMES = [
    ("crash", "2026-07-17", "2026-08-01", "暴落・韓国連動"),
    ("range", "2026-08-03", "2026-12-31", "レンジ回復"),
]

# イベントカレンダー(手動テーブル)。dateはJST。impact=翌営業日にも波及フラグ
EVENTS = [
    ("2026-07-31", "285A決算"),
    ("2026-08-08", "米雇用統計"),
    ("2026-08-12", "米CPI"),
    ("2026-08-13", "SNDK Investor Day"),
]

# walk-forward: 営業日の前半でパラメータ探索 → 後半OOS検証
MIN_TRADES_FOR_CONCLUSION = 10   # n<10 は「参考」表記

# ---------------------------------------------------------------- パス
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_RAW = os.path.join(BASE_DIR, "data_raw")
OUTPUT = os.path.join(BASE_DIR, "output")
