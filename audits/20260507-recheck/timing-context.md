# 計測時刻帯コンテキスト

| baseline | 時刻帯（home job 開始） | UTC | 備考 |
|---|---|---|---|
| 4/21 | 2026-04-21 **10:16 JST**（午前） | 2026-04-21T01:16:01Z | `audits/20260421/lighthouse-mobile/home.json` の `fetchTime` |
| 5/5 | 2026-05-05 **09:51 JST**（午前 dry-run）→ 16:13/16:42 JST（午後 batches） | 2026-05-05T00:51Z 〜 2026-05-06T07:45Z | dry-run と batch 3 の間に大きな時間差あり、計測時刻帯は混在 |
| 5/7 recheck（本タスク） | 2026-05-07 **02:26 JST**（深夜） | 2026-05-06T17:26Z | run-1 開始予定 |

## 同時刻帯条件の評価

- 4/21（午前）vs 5/7（深夜）= **時刻帯不一致**
- 5/5（午後混在）vs 5/7（深夜）= **時刻帯不一致**
- ネットワーク経路や Vercel Edge の負荷状態が時刻によって変わる可能性は完全に排除できない
- ただし、本 recheck は **同一 5/7 セッション内で 3 ラン連続実行**するため、3 ランの **内部分散** は時刻帯影響を受けない（同じ深夜帯で揃う）
- 仮に 5/7 中央値が 4/21 と乖離した場合、それが「真の回帰」か「時刻帯由来」かを切り分けるには 4/21 と同じ午前帯で追加 1 ランが必要（Case C 想定）

## 現環境

- OS: Windows 11 / Git Bash
- Node v24.14.0 / Lighthouse 13.2.0（npx --yes、npm cache 経由）
- Chrome: `C:\Program Files\Google\Chrome\Application\chrome.exe`
- 4/21 計測スクリプトの options を完全踏襲: `--form-factor=mobile --output=json --chrome-flags="--headless --no-sandbox" --quiet --max-wait-for-load=60000`
