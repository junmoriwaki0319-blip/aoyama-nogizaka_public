#!/bin/bash
set -u
OUT_DIR="./audits/20260421/axe"
LOG="./audits/20260421/axe/_run.log"
mkdir -p "$OUT_DIR"
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
  out="$OUT_DIR/$slug.json"
  echo "[$(date +%H:%M:%S)] START axe $slug" | tee -a "$LOG"
  npx --yes @axe-core/cli "$url" --save "$out" >> "$LOG" 2>&1
  rc=$?
  if [ -f "$out" ]; then
    echo "[$(date +%H:%M:%S)] OK-ish $slug (rc=$rc, size=$(wc -c < "$out") bytes)" | tee -a "$LOG"
  else
    echo "[$(date +%H:%M:%S)] FAIL $slug (rc=$rc)" | tee -a "$LOG"
  fi
done
echo "[$(date +%H:%M:%S)] DONE axe" | tee -a "$LOG"
