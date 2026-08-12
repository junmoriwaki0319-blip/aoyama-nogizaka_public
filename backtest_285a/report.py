# -*- coding: utf-8 -*-
"""出力生成: results.csv / daily_replay/{date}.md / report.md"""
import csv
import os

import config
import backtest


def _fmt(x, digits=0):
    if x is None:
        return "-"
    return f"{x:,.{digits}f}"


def _stat_row(name, s):
    ref = " ※参考(n<10)" if s["reference_only"] else ""
    return (f"| {name} | {s['n_days']} | {s['n_cards']} | {s['n_trades']} | "
            f"{s['trigger_rate']*100:.1f}% | {s['win_rate']*100:.1f}% | "
            f"{_fmt(s['rr'],2)} | {_fmt(s['expectancy'])}円 | {_fmt(s['total_pnl'])}円{ref} |")


_STAT_HEADER = ("| 条件 | 営業日 | カード数 | 発動数 | 発動率 | 勝率 | 平均RR | 期待値/枚 | 総損益 |\n"
                "|---|---|---|---|---|---|---|---|---|")


def write_results_csv(full_results, path):
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["date", "card", "side", "triggered", "entry_level", "fill_ts",
                    "fill_price", "exit_ts", "exit_price", "exit_reason",
                    "pnl_yen", "mae_yen", "mfe_yen", "auction_slip_yen"])
        for day in full_results:
            for c in day.cards:
                trig = c.status == "closed" or c.fill_ts is not None
                w.writerow([
                    day.date.isoformat(), c.kind, c.side, int(trig),
                    f"{c.entry_level:.0f}" if c.entry_level else "",
                    c.fill_ts.strftime("%H:%M") if c.fill_ts else "",
                    f"{c.fill_price:.0f}" if c.fill_price else "",
                    c.exit_ts.strftime("%H:%M") if c.exit_ts else "",
                    f"{c.exit_price:.0f}" if c.exit_price else "",
                    c.exit_reason,
                    f"{c.pnl:.0f}" if c.pnl is not None else "",
                    f"{c.mae:.0f}" if c.mae is not None else "",
                    f"{c.mfe:.0f}" if c.mfe is not None else "",
                    f"{day.auction_slip:.1f}" if day.auction_slip is not None else "",
                ])


def write_daily_replays(full_results, dirpath, synthetic=False):
    os.makedirs(dirpath, exist_ok=True)
    for day in full_results:
        snap = day.snapshot
        lines = []
        if synthetic:
            lines.append("> ⚠️ **SYNTHETIC DATA** — 合成データによるパイプライン検証用サンプル。実相場ではない。\n")
        lines.append(f"# {day.date} デイリーリプレイ (285A.T)\n")
        lines.append("## 8:55 スナップショット\n")
        lines.append(f"- 前日 高値/安値/終値: {_fmt(snap.prev_high)} / {_fmt(snap.prev_low)} / {_fmt(snap.prev_close)}")
        lines.append(f"- 5日 高値/安値: {_fmt(snap.hi5)} / {_fmt(snap.lo5)}")
        lines.append(f"- MA5/25/75: {_fmt(snap.ma5)} / {_fmt(snap.ma25)} / {_fmt(snap.ma75)}")
        lines.append(f"- 値幅制限: {_fmt(snap.limit_low)} 〜 {_fmt(snap.limit_high)}")
        if snap.vp_bands:
            bands = ", ".join(f"{lo:.0f}-{hi:.0f}({sh*100:.0f}%)" for lo, hi, sh in snap.vp_bands)
            lines.append(f"- しこり帯(過去20日 価格帯別出来高 上位): {bands}")
        if snap.gaps:
            gs = ", ".join(f"{lo:.0f}-{hi:.0f}({d})" for lo, hi, d in snap.gaps[-3:])
            lines.append(f"- 未埋め窓: {gs}")
        if snap.adr_theoretical:
            lines.append(f"- ADR理論値: {_fmt(snap.adr_theoretical)} (USDJPY={_fmt(snap.usdjpy_855,2)}, "
                         f"予測ギャップ {snap.adr_gap_pred:+.0f}円)")
        if snap.niy_855:
            lines.append(f"- NIY=F(8:55): {_fmt(snap.niy_855)}")
        lines.append(f"- {snap.notes[0]}")
        lines.append("\n## 生成カード\n")
        lines.append("| 型 | 方向 | 座標 | 損切り | 利確 | 備考 | 結果 |")
        lines.append("|---|---|---|---|---|---|---|")
        for c in day.cards:
            result = "未発動"
            if c.fill_ts is not None:
                result = f"約定{c.fill_price:.0f}→{c.exit_reason} {c.pnl:+.0f}円" if c.pnl is not None else "約定"
            lines.append(f"| {c.kind} | {'買' if c.side=='long' else '売'} | "
                         f"{_fmt(c.entry_level)} | {_fmt(c.stop)} | {_fmt(c.target)} | {c.note} | {result} |")
        lines.append("\n## 執行ログ\n")
        if day.log:
            lines.extend(f"- {x}" for x in day.log)
        else:
            lines.append("- (発動なし)")
        if day.kr_blocked_from:
            lines.append(f"- 韓国リスクオフ発動: {day.kr_blocked_from:%H:%M} 以降新規停止")
        if day.auction_slip is not None:
            lines.append(f"- 引けオークション滑り(日足終値−最終5分バー): {day.auction_slip:+.1f}円")
        with open(os.path.join(dirpath, f"{day.date}.md"), "w") as f:
            f.write("\n".join(lines) + "\n")


def _suggestions(out) -> list:
    """計測結果から「現行ルールから変えるべき点」上位5つをデータ駆動で作る。"""
    sugg = []
    bp = out["best_params"]

    # 1) 損切り幅
    base_rows = [g for g in out["grid"] if g["decay_n"] == bp["decay_n"]
                 and g["time_rule"] == bp["time_rule"]]
    cur = next((g for g in base_rows if g["stop"] == config.STOP_WIDTH_BASE), None)
    best_sw = max(base_rows, key=lambda g: g["is"]["expectancy"]) if base_rows else None
    if cur and best_sw and best_sw["stop"] != config.STOP_WIDTH_BASE:
        diff = best_sw["is"]["expectancy"] - cur["is"]["expectancy"]
        sugg.append((abs(diff),
                     f"損切り幅を{config.STOP_WIDTH_BASE:.0f}円→{best_sw['stop']:.0f}円へ変更 "
                     f"(IS期待値 {diff:+.0f}円/枚, n={best_sw['is']['n_trades']}"
                     f"{'・参考' if best_sw['is']['reference_only'] else ''})"))

    # 2) 減衰N
    n_rows = [g for g in out["grid"] if g["stop"] == bp["stop"] and g["time_rule"] == bp["time_rule"]]
    cur_n = next((g for g in n_rows if g["decay_n"] == config.DECAY_N_BASE), None)
    best_n = max(n_rows, key=lambda g: g["is"]["expectancy"]) if n_rows else None
    if cur_n and best_n and best_n["decay_n"] != config.DECAY_N_BASE:
        diff = best_n["is"]["expectancy"] - cur_n["is"]["expectancy"]
        sugg.append((abs(diff),
                     f"減衰印字をN={config.DECAY_N_BASE}→N={best_n['decay_n']}へ変更 "
                     f"(IS期待値 {diff:+.0f}円/枚, n={best_n['is']['n_trades']}"
                     f"{'・参考' if best_n['is']['reference_only'] else ''})"))

    # 3) 時間規制
    tr_rows = [g for g in out["grid"] if g["stop"] == bp["stop"] and g["decay_n"] == bp["decay_n"]]
    if len(tr_rows) == 2:
        on = next(g for g in tr_rows if g["time_rule"])
        off = next(g for g in tr_rows if not g["time_rule"])
        diff = off["is"]["expectancy"] - on["is"]["expectancy"]
        better, worse = ("OFF(大引けまで)", "ON") if diff > 0 else ("ON(14:30/15:20)", "OFF")
        sugg.append((abs(diff),
                     f"時間規制は{better}が優位 ({abs(diff):.0f}円/枚差, "
                     f"n={max(on['is']['n_trades'], off['is']['n_trades'])}"
                     f"{'・参考' if on['is']['reference_only'] else ''})"))

    # 4) ADRアンカー
    a_on, a_off = out["timing_adr"]["on"], out["timing_adr"]["off"]
    diff = a_on["expectancy"] - a_off["expectancy"]
    verb = "維持(座標調整に有効)" if diff > 0 else "廃止検討(改善が確認できない)"
    sugg.append((abs(diff),
                 f"ADRアンカーは{verb} (期待値差 {diff:+.0f}円/枚, n={a_on['n_trades']}"
                 f"{'・参考' if a_on['reference_only'] else ''})"))

    # 5) 韓国データ
    kr = out["timing_kr"]
    diff_rt = kr["realtime"]["expectancy"] - kr["off"]["expectancy"]
    diff_dl = kr["delayed"]["expectancy"] - kr["off"]["expectancy"]
    sugg.append((abs(diff_rt),
                 f"韓国リスクオフ・フィルター: リアルタイム{diff_rt:+.0f}円/枚, "
                 f"20分遅延{diff_dl:+.0f}円/枚 (対 不使用, n={kr['realtime']['n_trades']}"
                 f"{'・参考' if kr['realtime']['reference_only'] else ''})"))

    # 6) カード型別: 期待値が最悪の型の停止検討
    ck = {k: v for k, v in out["card_stats"].items() if v["n_trades"] > 0}
    if ck:
        worst = min(ck.items(), key=lambda kv: kv[1]["expectancy"])
        if worst[1]["expectancy"] < 0:
            sugg.append((abs(worst[1]["expectancy"]),
                         f"{worst[0]}型カードの停止/条件強化を検討 "
                         f"(期待値 {worst[1]['expectancy']:+.0f}円/枚, n={worst[1]['n_trades']}"
                         f"{'・参考' if worst[1]['reference_only'] else ''})"))
        best_k = max(ck.items(), key=lambda kv: kv[1]["expectancy"])
        if best_k[1]["expectancy"] > 0 and best_k[0] != worst[0]:
            sugg.append((best_k[1]["expectancy"] * 0.5,
                         f"{best_k[0]}型カードへの配分増を検討 "
                         f"(期待値 {best_k[1]['expectancy']:+.0f}円/枚, n={best_k[1]['n_trades']}"
                         f"{'・参考' if best_k[1]['reference_only'] else ''})"))

    # 7) イベント日
    ev, nm = out["event_stats"]["event"], out["event_stats"]["normal"]
    if ev["n_trades"] > 0:
        diff = ev["expectancy"] - nm["expectancy"]
        if diff < 0:
            sugg.append((abs(diff),
                         f"イベント日(決算・米指標)は新規カードを停止/縮小 "
                         f"(平常日比 {diff:+.0f}円/枚, n={ev['n_trades']}"
                         f"{'・参考' if ev['reference_only'] else ''})"))

    # n<10(「参考」)の提案は重みを1/4に減衰させ、確度の高い提案を上位に
    sugg.sort(key=lambda x: -(x[0] * (0.25 if "参考" in x[1] else 1.0)))
    return [s for _, s in sugg[:5]]


def write_report(out, path, synthetic=False):
    L = []
    if synthetic:
        L.append("> ⚠️ **SYNTHETIC DATA** — 本レポートは合成データによるパイプライン動作検証サンプルです。"
                 "数値は実相場の成績ではありません。実データ取得後に `./run_all.sh` で再生成してください。\n")
    L.append("# 285A.T デイトレカード・バックテスト結果\n")
    sessions = out["sessions"]
    L.append(f"- 対象期間: {sessions[0]} 〜 {sessions[-1]} ({len(sessions)}営業日)")
    L.append(f"- walk-forward: IS {out['is_days'][0]}〜{out['is_days'][-1]} ({len(out['is_days'])}日) / "
             f"OOS {out['oos_days'][0]}〜{out['oos_days'][-1]} ({len(out['oos_days'])}日)")
    bp = out["best_params"]
    L.append(f"- IS選択パラメータ: 損切り幅={bp['stop']:.0f}円, 減衰N={bp['decay_n']}, "
             f"時間規制={'ON' if bp['time_rule'] else 'OFF'}")
    L.append("- 絶対原則: ルックアヘッド禁止(8:55スナップショット、場中は5分足のみ)。"
             "n<10の結論は全て「参考」表記。")
    L.append("- 除外情報: 引け残り板・PTS・群衆情報(過去再現不能のため不使用)。")
    slip = out["auction_slip"]
    if slip["mean"] is not None:
        L.append(f"- 引けオークション滑り(日足終値−最終5分バー): 平均 {slip['mean']:+.1f}円 / "
                 f"平均絶対値 {slip['abs_mean']:.1f}円 (n={slip['n']})。"
                 "5分足は引けオークション非捕捉のため、大引け手仕舞いの実効コストとして考慮。")

    L.append("\n## OOS検証(アウトオブサンプル)\n")
    L.append(_STAT_HEADER)
    L.append(_stat_row("OOS(後半)", out["oos_stats"]))

    L.append("\n## カード型別(全期間・基準パラメータ)\n")
    L.append(_STAT_HEADER)
    labels = {"A": "A 押し目買い", "B": "B 戻り売り", "C": "C 深部受け", "D": "D 突破"}
    for k in ("A", "B", "C", "D"):
        L.append(_stat_row(labels[k], out["card_stats"][k]))

    L.append("\n## レジーム別(全期間混合の平均は出さない)\n")
    L.append(_STAT_HEADER)
    for key, v in out["regime_stats"].items():
        L.append(_stat_row(f"{v['label']} ({v['range'][0]}〜{v['range'][1]})", v["stats"]))

    L.append("\n## イベント層別\n")
    L.append(_STAT_HEADER)
    for cls, label in [("event", "イベント日"), ("event+1", "イベント翌営業日"), ("normal", "平常日")]:
        L.append(_stat_row(label, out["event_stats"][cls]))
    ev_list = ", ".join(f"{d}({name})" for d, name in config.EVENTS)
    L.append(f"\nイベントカレンダー(手動): {ev_list}")

    L.append("\n## パラメータ感度(IS区間グリッド)\n")
    L.append("| 損切り幅 | 減衰N | 時間規制 | 発動数 | 勝率 | 期待値/枚 | 総損益 |")
    L.append("|---|---|---|---|---|---|---|")
    for g in out["grid"]:
        s = g["is"]
        ref = " ※参考" if s["reference_only"] else ""
        L.append(f"| {g['stop']:.0f} | {g['decay_n']} | {'ON' if g['time_rule'] else 'OFF'} | "
                 f"{s['n_trades']} | {s['win_rate']*100:.0f}% | {s['expectancy']:,.0f}円 | "
                 f"{s['total_pnl']:,.0f}円{ref} |")

    L.append("\n## 情報タイミング感度(どの材料をどのタイミングで取ると何bp改善するか)\n")
    L.append("bp換算は約定価格平均に対する期待値差(1bp≈0.01%)。\n")
    L.append(_STAT_HEADER)
    L.append(_stat_row("(a) ADRアンカーあり", out["timing_adr"]["on"]))
    L.append(_stat_row("(a) ADRアンカーなし", out["timing_adr"]["off"]))
    for mode, label in [("realtime", "(b) 韓国RT"), ("delayed", "(b) 韓国20分遅延"), ("off", "(b) 韓国不使用")]:
        L.append(_stat_row(label, out["timing_kr"][mode]))

    # 期待値差とbp換算(1株あたり期待値差 ÷ 平均約定価格)
    trades = [t for r in out["full_results"] for t in r.trades if t.fill_price]
    avg_px = (sum(t.fill_price for t in trades) / len(trades)) if trades else 0.0

    def _bp(diff):
        if not avg_px:
            return "-"
        return f"{diff / config.LOT / avg_px * 10000:+.1f}bp"

    a_diff = out["timing_adr"]["on"]["expectancy"] - out["timing_adr"]["off"]["expectancy"]
    kr = out["timing_kr"]
    kr_rt = kr["realtime"]["expectancy"] - kr["off"]["expectancy"]
    kr_dl = kr["delayed"]["expectancy"] - kr["off"]["expectancy"]
    L.append("\n| 材料 | タイミング | 期待値差(円/枚) | bp換算(対平均約定価格) |")
    L.append("|---|---|---|---|")
    L.append(f"| ADR理論値 | 8:55(寄り前) | {a_diff:+,.0f} | {_bp(a_diff)} |")
    L.append(f"| 韓国(000660) | リアルタイム | {kr_rt:+,.0f} | {_bp(kr_rt)} |")
    L.append(f"| 韓国(000660) | 20分遅延 | {kr_dl:+,.0f} | {_bp(kr_dl)} |")

    L.append("\n## 現行ルールから変えるべき点(上位5・根拠n数と改善幅つき)\n")
    for i, s in enumerate(_suggestions(out), 1):
        L.append(f"{i}. {s}")

    L.append("\n## モデリング上の選択(要検証事項)\n")
    L.append("- 同時保有は1枚(2枚目は1枚目クローズ後に発動可)。")
    L.append("- C/Dの成行エントリーはシグナルバー終値+1tickで近似(次バー寄り相当)。")
    L.append("- ADRアンカーON時は予測ギャップ×0.5をカード座標にシフト(1%未満は無視)。")
    L.append("- 韓国リスクオフ: 000660.KSが寄比-2%で当日新規停止。")
    L.append("- スリッページ: 成行・逆指値に1tick(10円)不利適用。指値は指値約定。")

    with open(path, "w") as f:
        f.write("\n".join(L) + "\n")


def generate_all(out, output_dir, synthetic=False):
    os.makedirs(output_dir, exist_ok=True)
    write_results_csv(out["full_results"], os.path.join(output_dir, "results.csv"))
    write_daily_replays(out["full_results"], os.path.join(output_dir, "daily_replay"),
                        synthetic=synthetic)
    write_report(out, os.path.join(output_dir, "report.md"), synthetic=synthetic)
