#!/bin/bash
# 全期間の欠損データを5バッチに分割して順次再パース
# 使い方: EDINET_API_KEY=xxx bash scripts/run_reparse_all_batches.sh
set -e

cd "$(dirname "$0")/.."

echo "=== 全期間再パース（5バッチ順次実行） ==="
echo "開始: $(date)"
echo ""

# Batch 1: 2026-01〜2026-03 (already running separately, skip if progress exists)
echo "--- Batch 1: 2026-01 to 2026-03 ---"
node scripts/reparse_all.js --from 2026-01 --to 2026-03
echo ""

# Batch 2: 2025-09〜2025-12
echo "--- Batch 2: 2025-09 to 2025-12 ---"
node scripts/reparse_all.js --from 2025-09 --to 2025-12
echo ""

# Batch 3: 2025-05〜2025-08
echo "--- Batch 3: 2025-05 to 2025-08 ---"
node scripts/reparse_all.js --from 2025-05 --to 2025-08
echo ""

# Batch 4: 2025-01〜2025-04
echo "--- Batch 4: 2025-01 to 2025-04 ---"
node scripts/reparse_all.js --from 2025-01 --to 2025-04
echo ""

# Batch 5: 2024-09〜2024-12
echo "--- Batch 5: 2024-09 to 2024-12 ---"
node scripts/reparse_all.js --from 2024-09 --to 2024-12
echo ""

echo "=== 全バッチ完了: $(date) ==="

# 完了後 git push
echo "Committing and pushing..."
git add data/reports.json
git commit -m "data: 全期間の大量保有報告データ再パース（保有比率・目的の欠損補完）"
git push origin main
echo "Done!"
