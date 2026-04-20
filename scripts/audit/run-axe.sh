#!/bin/bash
set -u
# T2 axe-core CLI runner for 13 URLs

OUT_DIR="./audits/20260420"
mkdir -p "$OUT_DIR"
LOG="$OUT_DIR/axe.log"
: > "$LOG"

urls=(
  "home|https://aoyama-nogizaka.com/"
  "team|https://aoyama-nogizaka.com/team"
  "news|https://aoyama-nogizaka.com/news/"
  "privacy|https://aoyama-nogizaka.com/privacy"
  "activist-dashboard|https://aoyama-nogizaka.com/activist-dashboard.html"
  "risk-assessment|https://aoyama-nogizaka.com/risk-assessment.html"
  "activist-screener|https://aoyama-nogizaka.com/activist-screener.html"
  "food-service|https://aoyama-nogizaka.com/food-service.html"
  "saas|https://aoyama-nogizaka.com/saas.html"
  "ad-agency|https://aoyama-nogizaka.com/ad-agency.html"
  "digital-media|https://aoyama-nogizaka.com/digital-media.html"
  "entertainment-sector-dashboard|https://aoyama-nogizaka.com/entertainment-sector-dashboard.html"
  "news-activist-shareholder-proposals-japan|https://aoyama-nogizaka.com/news/activist-shareholder-proposals-japan.html"
)

for entry in "${urls[@]}"; do
  slug="${entry%%|*}"
  url="${entry#*|}"
  out="$OUT_DIR/axe-$slug.json"
  echo "[$(date +%H:%M:%S)] START axe $slug" | tee -a "$LOG"
  # Skip home (already done)
  if [ "$slug" = "home" ] && [ -f "$out" ]; then
    echo "[$(date +%H:%M:%S)] SKIP axe $slug (already exists)" | tee -a "$LOG"
    continue
  fi
  npx --yes @axe-core/cli "$url" --save "$out" >> "$LOG" 2>&1
  rc=$?
  if [ -f "$out" ]; then
    echo "[$(date +%H:%M:%S)] OK-ish axe $slug (rc=$rc, size=$(wc -c < "$out") bytes)" | tee -a "$LOG"
  else
    echo "[$(date +%H:%M:%S)] FAIL axe $slug (rc=$rc)" | tee -a "$LOG"
  fi
done
echo "[$(date +%H:%M:%S)] DONE axe" | tee -a "$LOG"
