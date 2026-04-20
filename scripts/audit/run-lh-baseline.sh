#!/bin/bash
set -u
# T1: Lighthouse mobile baseline runner for 13 URLs

OUT_DIR="./audits/20260420"
mkdir -p "$OUT_DIR"
LOG="$OUT_DIR/lh-baseline.log"
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
  out="$OUT_DIR/lh-mobile-$slug.json"
  echo "[$(date +%H:%M:%S)] START $slug -> $url" | tee -a "$LOG"
  npx --yes lighthouse "$url" \
    --form-factor=mobile \
    --output=json \
    --output-path="$out" \
    --chrome-flags="--headless --no-sandbox" \
    --quiet \
    --max-wait-for-load=60000 \
    >> "$LOG" 2>&1
  rc=$?
  if [ $rc -eq 0 ] && [ -f "$out" ]; then
    echo "[$(date +%H:%M:%S)] OK $slug (rc=$rc, size=$(wc -c < "$out") bytes)" | tee -a "$LOG"
  else
    echo "[$(date +%H:%M:%S)] FAIL $slug (rc=$rc)" | tee -a "$LOG"
  fi
done
echo "[$(date +%H:%M:%S)] DONE T1" | tee -a "$LOG"
