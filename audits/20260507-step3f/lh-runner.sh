#!/bin/bash
set -uo pipefail
PREVIEW_URL="$1"
OUT_DIR="audits/20260507-step3f/lh-base421-on-current-infra"
mkdir -p "$OUT_DIR"
RUN_LOG="$OUT_DIR/_run.log"
: > "$RUN_LOG"

declare -a entries=(
  "home:"
  "activist-dashboard:activist-dashboard.html"
  "risk-assessment:risk-assessment.html"
  "activist-screener:activist-screener.html"
  "team:team"
  "news:news/"
  "news-activist-report:news/activist-report.html"
  "food-service:food-service.html"
  "saas:saas.html"
  "ad-agency:ad-agency.html"
  "digital-media:digital-media.html"
  "entertainment:entertainment-sector-dashboard.html"
  "privacy:privacy"
)

for entry in "${entries[@]}"; do
  fname="${entry%%:*}"
  page="${entry#*:}"
  url="${PREVIEW_URL%/}/${page}"
  out="$OUT_DIR/lh-mobile-${fname}.json"
  echo "[$(date +%H:%M:%S)] START $fname -> $url" >> "$RUN_LOG"
  npx --yes lighthouse "$url" \
    --form-factor=mobile \
    --throttling-method=simulate \
    --output=json \
    --output-path="$out" \
    --chrome-flags="--headless --no-sandbox" \
    >> "$RUN_LOG" 2>&1
  rc=$?
  size=0
  [ -f "$out" ] && size=$(wc -c < "$out")
  echo "[$(date +%H:%M:%S)] DONE $fname rc=$rc size=$size" >> "$RUN_LOG"
  git add "$out" 2>/dev/null
done
git add "$RUN_LOG" 2>/dev/null
echo "all-done"
