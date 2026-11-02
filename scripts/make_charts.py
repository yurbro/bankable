#!/usr/bin/env python3
"""
make_charts.py — 从 model_outputs.json 生成商业计划书用的财务图表（PNG）。

用法:
    python make_charts.py model_outputs.json -o charts/ [--lang zh|en] [--dpi 200]

生成（视数据而定）:
    revenue_profit.png   营业收入与净利润（柱）+ 毛利率标注
    cash_balance.png     期末现金与最低现金要求
    scenarios.png        三情景的收入与期末现金
    use_of_funds.png     资金用途（需 outputs 中有 use_of_funds）
    dscr.png             偿债覆盖率（有借款时）
    monthly_cash.png     第 1 年月度现金（有月度表时）

在 plan.md 中引用：![图 3 收入与净利润预测](charts/revenue_profit.png)
"""
import argparse
import json
import os

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402

PALETTE = {"blue": "#2a78d6", "orange": "#eb6834", "aqua": "#1baf7a", "yellow": "#eda100", "gray": "#9a9892", "ink": "#0b0b0b", "muted": "#52514e"}
SERIES = [PALETTE["blue"], PALETTE["orange"], PALETTE["aqua"], PALETTE["yellow"]]

LABELS = {
    "zh": {
        "revenue": "营业收入", "net_income": "净利润", "gross_margin": "毛利率", "closing_cash": "期末现金", "min_cash": "最低现金要求",
        "base": "基准", "downside": "下行", "upside": "上行", "unit": "（{unit}）", "year": "年份",
        "revenue_profit_title": "营业收入与净利润预测", "cash_title": "期末现金预测", "scen_rev_title": "三情景营业收入", "scen_cash_title": "三情景期末现金",
        "uof_title": "资金用途", "dscr_title": "偿债覆盖率（DSCR）", "dscr_ref": "银行常见要求 1.2 / 1.5", "monthly_title": "第 1 年月度期末现金", "month": "月",
        "source": "来源：财务模型（基准情景）",
    },
    "en": {
        "revenue": "Revenue", "net_income": "Net income", "gross_margin": "Gross margin", "closing_cash": "Closing cash", "min_cash": "Minimum cash",
        "base": "Base", "downside": "Downside", "upside": "Upside", "unit": " ({unit})", "year": "Year",
        "revenue_profit_title": "Revenue and net income forecast", "cash_title": "Closing cash forecast", "scen_rev_title": "Revenue by scenario", "scen_cash_title": "Closing cash by scenario",
        "uof_title": "Use of funds", "dscr_title": "Debt service coverage ratio (DSCR)", "dscr_ref": "Typical bank covenant 1.2 / 1.5", "monthly_title": "Year-1 monthly closing cash", "month": "Month",
        "source": "Source: financial model (base case)",
    },
}
UNIT_EN = {"万元": "RMB 10k", "亿元": "RMB 100m", "元": "RMB", "CNY": "CNY", "USD": "USD"}


def setup_fonts(lang):
    candidates = ["Noto Sans CJK SC", "Noto Sans CJK JP", "Source Han Sans SC", "Microsoft YaHei", "PingFang SC", "SimHei", "WenQuanYi Micro Hei", "Arial Unicode MS"]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for c in candidates:
        if c in available:
            plt.rcParams["font.family"] = [c, "DejaVu Sans"]
            break
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams.update({
        "font.size": 10, "axes.titlesize": 11, "axes.titleweight": "bold", "axes.edgecolor": "#cfcfcb", "axes.linewidth": 0.8,
        "axes.spines.top": False, "axes.spines.right": False, "axes.grid": True, "grid.color": "#e6e6e2", "grid.linewidth": 0.6,
        "axes.axisbelow": True, "xtick.color": PALETTE["muted"], "ytick.color": PALETTE["muted"], "legend.frameon": False, "figure.dpi": 100,
    })


def fmt_num(v):
    if v is None:
        return ""
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    return f"{v:,.1f}"


def save(fig, path, dpi, note=None):
    if note:
        fig.text(0.01, 0.01, note, fontsize=7.5, color=PALETTE["muted"], ha="left", va="bottom")
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(path, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("[OK]", path)


def chart_revenue_profit(o, L, unit, out, dpi):
    s = o["scenarios"]["base"]["series"]
    years = [str(y) for y in o["years"]]
    rev, ni, gm = s["revenue"], s["net_income"], s.get("gross_margin", [])
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    x = range(len(years))
    w = 0.38
    ax.bar([i - w / 2 for i in x], rev, width=w, color=SERIES[0], label=L["revenue"])
    ax.bar([i + w / 2 for i in x], ni, width=w, color=SERIES[1], label=L["net_income"])
    for i, v in enumerate(rev):
        if v:
            ax.text(i - w / 2, v, fmt_num(v), ha="center", va="bottom", fontsize=8, color=PALETTE["ink"])
    for i, v in enumerate(ni):
        if v:
            ax.text(i + w / 2, v if v >= 0 else 0, fmt_num(v), ha="center", va="top" if v < 0 else "bottom", fontsize=8, color=PALETTE["ink"])
    if gm and any(g is not None for g in gm):
        for i, g in enumerate(gm):
            if g is not None and rev[i]:
                ax.text(i, max(rev[i], 0) * 1.08, f"{L['gross_margin']} {g * 100:.0f}%", ha="center", va="bottom", fontsize=7.5, color=PALETTE["muted"])
        ax.set_ylim(top=max(rev) * 1.25)
        if min(ni) < 0:
            ax.set_ylim(bottom=min(ni) * 1.3)
    ax.set_xticks(list(x))
    ax.set_xticklabels(years)
    ax.axhline(0, color="#999", linewidth=0.8)
    ax.set_title(L["revenue_profit_title"] + L["unit"].format(unit=unit))
    ax.legend(loc="upper left")
    save(fig, os.path.join(out, "revenue_profit.png"), dpi, L["source"])


def chart_cash(o, L, unit, out, dpi):
    s = o["scenarios"]["base"]["series"]
    years = [str(y) for y in o["years"]]
    cash = s["closing_cash"]
    mc = o.get("min_cash_requirement", 0) or 0
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.plot(years, cash, color=SERIES[0], linewidth=2, marker="o", markersize=5, label=L["closing_cash"])
    ax.fill_between(years, cash, 0, color=SERIES[0], alpha=0.08)
    for i, v in enumerate(cash):
        ax.text(i, v, fmt_num(v), ha="center", va="bottom", fontsize=8)
    if mc:
        ax.axhline(mc, color=PALETTE["orange"], linestyle="--", linewidth=1.2, label=f"{L['min_cash']} {fmt_num(mc)}")
    ax.axhline(0, color="#999", linewidth=0.8)
    ax.set_title(L["cash_title"] + L["unit"].format(unit=unit))
    ax.legend(loc="best")
    save(fig, os.path.join(out, "cash_balance.png"), dpi, L["source"])


def chart_scenarios(o, L, unit, out, dpi):
    years = [str(y) for y in o["years"]]
    codes = [c for c in ("base", "downside", "upside") if c in o["scenarios"]]
    if len(codes) < 2:
        return
    fig, axes = plt.subplots(1, 2, figsize=(9.6, 3.8))
    mc = o.get("min_cash_requirement", 0) or 0
    for j, (key, title) in enumerate([("revenue", L["scen_rev_title"]), ("closing_cash", L["scen_cash_title"])]):
        ax = axes[j]
        for k, code in enumerate(codes):
            ax.plot(years, o["scenarios"][code]["series"][key], color=SERIES[k], linewidth=2, marker="o", markersize=4, label=L[code])
        if key == "closing_cash":
            ax.axhline(0, color="#999", linewidth=0.8)
            if mc:
                ax.axhline(mc, color=PALETTE["gray"], linestyle="--", linewidth=1, label=L["min_cash"])
        ax.set_title(title + L["unit"].format(unit=unit))
        ax.legend(loc="best", fontsize=8)
    save(fig, os.path.join(out, "scenarios.png"), dpi, L["source"].replace("（基准情景）", "").replace(" (base case)", ""))


def chart_use_of_funds(o, L, unit, out, dpi):
    uof = o.get("use_of_funds")
    if not uof:
        return
    names = [u["name"] for u in uof]
    vals = [float(u["amount"]) for u in uof]
    total = sum(vals)
    fig, ax = plt.subplots(figsize=(7.2, 0.6 + 0.55 * len(names)))
    colors = [SERIES[i % len(SERIES)] if i < 4 else PALETTE["gray"] for i in range(len(names))]
    ax.barh(names[::-1], vals[::-1], color=colors[::-1], height=0.55)
    for i, v in enumerate(vals[::-1]):
        ax.text(v, i, f" {fmt_num(v)}（{v / total * 100:.0f}%）" if L is LABELS["zh"] else f" {fmt_num(v)} ({v / total * 100:.0f}%)", va="center", fontsize=8.5)
    ax.set_xlim(0, max(vals) * 1.35)
    ax.set_title(L["uof_title"] + L["unit"].format(unit=unit) + (f"，合计 {fmt_num(total)}" if L is LABELS["zh"] else f", total {fmt_num(total)}"))
    ax.grid(axis="y", visible=False)
    save(fig, os.path.join(out, "use_of_funds.png"), dpi)


def chart_dscr(o, L, out, dpi):
    d = o["scenarios"]["base"]["derived"].get("dscr_by_year") or []
    if not d:
        return
    years = [str(y) for y, _ in d]
    vals = [v for _, v in d]
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    colors = [SERIES[0] if v >= 1.2 else PALETTE["orange"] for v in vals]
    ax.bar(years, vals, color=colors, width=0.5)
    for i, v in enumerate(vals):
        ax.text(i, v, f"{v:.2f}", ha="center", va="bottom", fontsize=8.5)
    ax.axhline(1.2, color=PALETTE["gray"], linestyle="--", linewidth=1)
    ax.axhline(1.5, color=PALETTE["gray"], linestyle=":", linewidth=1)
    ax.text(len(years) - 0.5, 1.52, L["dscr_ref"], fontsize=7.5, color=PALETTE["muted"], ha="right", va="bottom")
    ax.set_ylim(bottom=min(0, min(vals) * 1.1))
    ax.set_title(L["dscr_title"])
    save(fig, os.path.join(out, "dscr.png"), dpi, L["source"])


def chart_monthly(o, L, unit, out, dpi):
    m = o["scenarios"]["base"].get("monthly_year1")
    if not m:
        return
    cash = m["closing_cash"]
    months = [f"{i + 1}{L['month']}" if L is LABELS["zh"] else f"M{i + 1}" for i in range(len(cash))]
    mc = o.get("min_cash_requirement", 0) or 0
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    ax.plot(months, cash, color=SERIES[0], linewidth=2, marker="o", markersize=4)
    ax.fill_between(months, cash, 0, color=SERIES[0], alpha=0.08)
    if mc:
        ax.axhline(mc, color=PALETTE["orange"], linestyle="--", linewidth=1.2, label=L["min_cash"])
        ax.legend(loc="best")
    ax.axhline(0, color="#999", linewidth=0.8)
    lo = min(cash)
    i_lo = cash.index(lo)
    ax.annotate(fmt_num(lo), (i_lo, lo), textcoords="offset points", xytext=(0, -14), ha="center", fontsize=8, color=PALETTE["orange"])
    ax.set_title(L["monthly_title"] + L["unit"].format(unit=unit) + f" — {o['years'][0]}")
    save(fig, os.path.join(out, "monthly_cash.png"), dpi, L["source"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("outputs")
    ap.add_argument("-o", "--outdir", default="charts")
    ap.add_argument("--lang", default="zh", choices=["zh", "en"])
    ap.add_argument("--dpi", type=int, default=200)
    args = ap.parse_args()
    with open(args.outputs, encoding="utf-8") as f:
        o = json.load(f)
    os.makedirs(args.outdir, exist_ok=True)
    setup_fonts(args.lang)
    L = LABELS[args.lang]
    unit = o.get("unit", "万元")
    if args.lang == "en":
        unit = UNIT_EN.get(unit, unit)
    chart_revenue_profit(o, L, unit, args.outdir, args.dpi)
    chart_cash(o, L, unit, args.outdir, args.dpi)
    chart_scenarios(o, L, unit, args.outdir, args.dpi)
    chart_use_of_funds(o, L, unit, args.outdir, args.dpi)
    chart_dscr(o, L, args.outdir, args.dpi)
    chart_monthly(o, L, unit, args.outdir, args.dpi)


if __name__ == "__main__":
    main()
