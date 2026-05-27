# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

Public corporate site for **青山乃木坂パートナーズ** (Aoyama Nogizaka Partners) — a Tokyo-based management consulting / M&A advisory firm. The site is primarily Japanese-language and combines a marketing site, several sector-research dashboards (activist screener/dashboard, sogo-shosha, semiconductor-trading, food-service, SaaS, ad-agency, digital-media, entertainment), and a few member-gated premium pages.

Production URL: `https://aoyama-nogizaka.com`

## Architecture

The repo is the document root itself — there is no build step for the HTML. Each top-level `*.html` file is the actual page served (e.g. `food-service.html` is reachable at both `/food-service.html` and `/food-service` via a `vercel.json` rewrite). Most `*.html` files are large self-contained files with their own inline `<style>` blocks; shared cross-cutting CSS lives in `css/` (notably `mobile-nav.css`).

Three deployment surfaces work together:

- **Firebase Hosting** (`firebase.json`) — primary hosting of the static files. Defines the enforcing `Content-Security-Policy` header. Apex domain via `CNAME`.
- **Vercel** (`vercel.json`) — hosts the serverless API under `api/` and provides `cleanUrls` rewrites for the static pages. CSP is `Report-Only` here (used as a staging/preview surface). Region: `hnd1`.
- **Firebase Cloud Functions** (`functions/`) — only one trigger: deleting the Firestore `users/{uid}` document when the Auth user is deleted.

Auth & data:

- **Firebase Auth + Firestore** is the membership layer for premium pages (activist-dashboard, activist-screener, food-service, saas, sogo-shosha, ad-agency, digital-media, entertainment, …). The Firestore rules (`firestore.rules`) allow each user to read/write only their own `users/{uid}` doc and allow any authenticated user to read `premiumContent/*`.
- The Firebase web SDK is loaded as inline ES modules from `gstatic.com` inside each gated `*.html`. Per-page wrappers live in `js/firebase-auth-*.js`. They expose `window.firebaseLogin`, `window.firebaseGoogleLogin`, etc., which `js/auth.js` then drives via event delegation on `data-auth-action` / `data-auth-submit` / `data-auth-change` attributes.
- The Vercel API verifies Firebase ID tokens against Google's public keys directly (see `api/premium-reports.js`), without using `firebase-admin`.

Data pipeline:

- `data/reports.json` — EDINET 大量保有報告書 (large-shareholding reports) ingested daily by `scripts/fetch_edinet.py` via the `update-edinet.yml` GitHub Action (runs at JST 08:00 / 16:00 / 19:00).
- `data/edinet-financials.json` — EDINET XBRL financials ingested by `scripts/fetch_edinet_financials.py` via `update-edinet-financials.yml` (daily 18:00 JST).
- `data/activist-campaigns.json` — built from `data/reports.json` + `scripts/known_activists.json` by `scripts/build-campaign-data.js`.
- `data/sogo-shosha/` — sector dashboard data (see `data/sogo-shosha/README.md`). `raw/` is Yahoo Finance + EDINET cache, `processed/` is the normalized output consumed by the page.
- `data/activist-materials/<campaign>/` — per-campaign Markdown/PDF research notes. Served as `text/plain` (header set in `vercel.json`).
- Vercel API routes under `api/edinet/`, `api/stock/`, `api/yahoo/`, plus `api/batch.js`, `api/stock-prices.js`, `api/premium-reports.js`.

The activist GitHub Action commits regenerated JSON directly to `main`, so expect frequent automated commits on that branch.

## Commands

```bash
# Install (root + Cloud Functions)
npm install
( cd functions && npm install )

# Playwright smoke tests — runs against PRODUCTION (baseURL=https://aoyama-nogizaka.com)
npx playwright test                              # all projects
npx playwright test --project=smoke              # unauthenticated smoke only
npx playwright test tests/smoke.spec.js          # single file
npx playwright test -g "ハンバーガーメニュー"     # single test by name

# Authenticated tests require TEST_EMAIL / TEST_PASSWORD env vars
TEST_EMAIL=... TEST_PASSWORD=... npx playwright test --project=authenticated

# Minify the shared CSS files (output: *.min.css)
npm run minify-css

# Regenerate CSP hashes after editing any inline JSON-LD <script> block
node scripts/generate-csp-hashes.js
# → then paste the new sha256-... values into firebase.json (script-src) and _headers

# Rebuild the activist campaigns dataset from reports.json
node scripts/build-campaign-data.js

# Manually run the EDINET ingest (needs EDINET_API_KEY)
EDINET_API_KEY=... python3 scripts/fetch_edinet.py
EDINET_API_KEY=... python3 scripts/fetch_edinet_financials.py
```

There is no project-level lint or typecheck step, and `npm test` is unconfigured (it intentionally exits 1) — use Playwright directly.

## Conventions

**Inline event handlers are forbidden.** The CSP rollout removed every `onclick` / `onsubmit` / `onchange` from this repo. There is a regression test that asserts `document.querySelectorAll('[onclick],[onsubmit],[onchange]').length === 0` on every key page (`tests/smoke.spec.js` → "inline handler 完全除去確認"). When adding interactivity:

- Add a `data-auth-action="openModal:login"` / `data-auth-submit="handleLogin"` style attribute and wire the handler in the page-specific JS under `js/` (already loaded by event delegation in `js/auth.js` for auth-related actions).
- New inline `<script>` (non-JSON-LD) is also off-limits under the enforcing CSP. Move logic into a file under `js/` and add `<script src="/js/...">`.
- Inline JSON-LD (`<script type="application/ld+json">`) is allowed because each block's sha256 is in the CSP. **After editing any JSON-LD block, rerun `node scripts/generate-csp-hashes.js` and update the hash list in `firebase.json` and `_headers`.** `CSP_DEPLOY_TODO.md` has the post-deploy checklist.

**Design system** — keep it tight; this is a buttoned-up financial-firm site, not a startup landing page. The full guide is in `.claude/skills/design-review/SKILL.md` (invoke via the `design-review` skill before touching visual styling). Short version:

- Colors are CSS variables: `--navy #1a2d4f`, `--gold #9b8b6e`, `--off-white #f8f7f5`, etc.
- Fonts: Noto Serif JP (headings), Cormorant Garamond (English/labels), Noto Sans JP (body). Don't introduce new families.
- Avoid the "AI slop" patterns called out in the skill: vivid gradients, huge hero text (>60px), heavy box-shadows, blob/wave SVG decoration, oversized border-radius (>8px), emoji.

**Page routing.** Don't link to `*.html` extensions in nav — Vercel rewrites bare paths (`/team`, `/food-service`, `/activist-dashboard`, …) to the HTML. The matching `<dir>/index.html` files (e.g. `team/index.html`) exist for Firebase Hosting's no-rewrite path.

**Premium API surface.** `api/premium-reports.js` verifies Firebase ID tokens against `https://www.googleapis.com/robot/v1/metadata/x509/securetoken@system.gserviceaccount.com`. It deliberately does not depend on `firebase-admin` (so it can run on Vercel's Node runtime without service-account credentials). Keep token verification logic consistent with that pattern if you add new gated endpoints.

**CSP changes.** If you add a new external host (script, font, API), update **both** `firebase.json` (enforcing CSP) and `vercel.json` / `_headers` (report-only CSP). The `connect-src` list in particular is easy to forget when adding a new third-party API.

**Auto-generated data files.** `data/reports.json`, `data/edinet-financials.json`, and `data/activist-campaigns.json` are regenerated by GitHub Actions on `main`. Don't hand-edit them — change the generating script or the curated input (`scripts/known_activists.json`) instead. Expect rebase conflicts on these files if your branch lives long.

**Git workflow.** Pre-commit hooks aren't configured; the Playwright tests run against production, not against a local server, so they are not a pre-merge gate. Verify visual changes manually (or via the `design-review` skill on the HTML you touched).
