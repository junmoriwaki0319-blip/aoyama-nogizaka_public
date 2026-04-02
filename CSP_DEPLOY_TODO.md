# Phase 5 CSP - デプロイ後 動作確認チェックリスト

## 作業完了状況
- [x] Phase 2: auth.js共通化 + data-auth-action委譲
- [x] Phase 3: 静的ページのonclick除去
- [x] Phase 4: activist-screener/dashboard等の全onclick除去
- [x] Phase 5: 全inline script外部化（18個のJSファイル）+ CSPヘッダー設定

## デプロイ後の確認ポイント

### 1. CSPヘッダー確認
- [ ] DevToolsのNetworkタブでレスポンスヘッダーにContent-Security-Policyが含まれること
- [ ] Consoleタブでscript-src/connect-src等のCSP違反エラーがないこと

### 2. ページ別動作確認
- [ ] **index.html** - お問い合わせフォーム送信
- [ ] **news/index.html** - フィルター切替（5カテゴリ）
- [ ] **news/activist-report.html** - モバイルナビ開閉・Chart.js表示・相談フォーム
- [ ] **risk-assessment.html** - データ取得・リスク評価・クリア
- [ ] **activist-dashboard.html** - タブ切替・フィルター・ソート・投資家詳細モーダル
- [ ] **activist-screener.html** - プリセット・ランキングソート・スキャン・認証モーダル
- [ ] **food-service.html** - DLボタン・認証モーダル開閉
- [ ] **saas.html** - DLボタン・認証モーダル開閉

### 3. 認証フロー（Junさん側で実施）
- [ ] Googleログイン
- [ ] マイページ表示
- [ ] ログアウト

### 4. CSP違反が見つかった場合
firebase.json / vercel.json のCSPヘッダーにドメイン追加が必要。
特に connect-src（API通信先）と script-src（外部スクリプト）を確認。

## 作成されたJSファイル一覧（/js/配下）
gtag.js, reveal.js, redirect.js, index-page.js, news-page.js,
activist-report-page.js, risk-assessment-page.js, activist-dashboard-page.js,
activist-screener-page.js, food-service-page.js, saas-page.js,
firebase-auth-dashboard.js, firebase-auth-screener.js,
firebase-auth-food-service.js, firebase-auth-saas.js,
nav-handler.js, auth.js, screener.js
