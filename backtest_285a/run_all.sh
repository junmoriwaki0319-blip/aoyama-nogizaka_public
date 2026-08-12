#!/usr/bin/env bash
# 285A.T デイトレカード・バックテスト: データ再取得〜レポート生成まで1コマンド
# 依存: Python 3.10+ (標準ライブラリのみ)
set -euo pipefail
cd "$(dirname "$0")"
python3 run_all.py "$@"
