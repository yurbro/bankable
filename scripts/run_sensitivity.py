#!/usr/bin/env python3
"""
run_sensitivity.py — 单因素敏感性分析：逐个改动关键假设，重算模型，汇总对关键结论的影响。

用法:
    python run_sensitivity.py assumptions.json -o work/sensitivity_outputs.json [--md work/sensitivity.md] [--chart work/charts/sensitivity.png] [--cases my_cases.json]

默认 12 个案例：销量 ±20%、价格 ±10%、单位成本 ±10%、费用 ±10%、资本开支 ±20%、应收账期 ±30 天。
自定义案例文件为数组，元素形如:
    {"id": "price-15", "label": "价格 −15%", "driver": "price", "mult": 0.85}          # driver ∈ volume/price/unit_cost/opex
    {"id": "capex+30", "label": "资本开支 +30%", "patch": "capex", "factor": 1.3}
    {"id": "ar+60",  "label": "应收账期 +60 天", "patch": "ar_days", "delta": 60}
    {"id": "tax15",  "label": "所得税 15%", "patch": "json", "path": "tax.income_tax_rate", "value": 0.15}
输出 JSON（基准 + 各案例的关键指标与相对基准的变化）、Markdown 表、龙卷风图（NPV 与最低期末现金）。
正文引用这些数字时，把 sensitivity_outputs.json 作为 check_consistency.py 的 --extra 传入。
"""
import argparse
import copy
import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_financial_model as bfm  # noqa: E402
import recalc_and_extract as rce  # noqa: E402

DEFAULT_CASES = [
    {"id": "volume-20", "label": "销量 −20%", "driver": "volume", "mult": 0.8, "pair": "销量 ±20%"},
    {"id": "volume+20", "label": "销量 +20%", "driver": "volume", "mult": 1.2, "pair": "销量 ±20%"},
    {"id": "price-10", "label": "价格 −10%", "driver": "price", "mult": 0.9, "pair": "价格 ±10%"},
    {"id": "price+10", "label": "价格 +10%", "driver": "price", "mult": 1.1, "pair": "价格 ±10%"},
    {"id": "unit_cost+10", "label": "单位成本 +10%", "driver": "unit_cost", "mult": 1.1, "pair": "单位成本 ±10%"},
    {"id": "unit_cost-10", "label": "单位成本 −10%", "driver": "unit_cost", "mult": 0.9, "pair": "单位成本 ±10%"},
    {"id": "opex+10", "label": "费用 +10%", "driver": "opex", "mult": 1.1, "pair": "费用 ±10%"},
    {"id": "opex-10", "label": "费用 −10%", "driver": "opex", "mult": 0.9, "pair": "费用 ±10%"},
    {"id": "capex+20", "label": "资本开支 +20%", "patch": "capex", "factor": 1.2, "pair": "资本开支 ±20%"},
    {"id": "capex-20", "label": "资本开支 −20%", "patch": "capex", "factor": 0.8, "pair": "资本开支 ±20%"},
    {"id": "ar+30", "label": "应收账期 +30 天", "patch": "ar_days", "delta": 30, "pair": "应收账期 ±30 天"},
    {"id": "ar-30", "label": "应收账期 −30 天", "patch": "ar_days", "delta": -30, "pair": "应收账期 ±30 天"},
]
METRICS = [
    ("revenue_last", "末年营业收入"), ("total_net_income", "累计净利润"), ("min_cash_balance", "最低期末现金"),
    ("funding_gap", "资金缺口"), ("peak_funding_need", "峰值资金需求"), ("irr_with_tv", "项目 IRR（含期末价值）"),
    ("npv", "NPV"), ("dscr_min", "最低 DSCR"), ("first_profit_year", "首个盈利年"),
]


def set_path(d, path, value):
    keys = path.split(".")
    cur = d
    for k in keys[:-1]:
        cur = cur[k]
    cur[keys[-1]] = value


def apply_case(a, case):
    a = copy.deepcopy(a)
    selector = 1
    if "driver" in case:
        a.setdefault("scenarios", {})["downside"] = {"volume": 1.0, "price": 1.0, "unit_cost": 1.0, "opex": 1.0, case["driver"]: float(case["mult"]), "label": case["label"]}
        selector = 2
    elif case.get("patch") == "capex":
        for c in a.get("capex", []) or []:
            c["amount"] = float(c["amount"]) * float(case["factor"])
    elif case.get("patch") == "ar_days":
        wc = a.setdefault("working_capital", {})
        wc["ar_days"] = max(0.0, float(wc.get("ar_days", 0)) + float(case["delta"]))
    elif case.get("patch") == "json":
        set_path(a, case["path"], case["value"])
    else:
        raise ValueError(f"无法识别的案例：{case}")
    return a, selector


def run_one(a, selector, soffice, tmpdir, tag, min_cash):
    d = os.path.join(tmpdir, tag)
    os.makedirs(d, exist_ok=True)
    model = os.path.join(d, "model.xlsx")
    bfm.build(a, model)
    import openpyxl
    wb = openpyxl.load_workbook(model)
    recs = rce.read_map(wb)
    sel = recs["scenario_selector"]
    wb["假设"].cell(row=sel["row"], column=sel["col"]).value = selector
    wb.save(model)
    calc = rce.recalc(soffice, model, os.path.join(d, "calc"))
    vals = rce.read_values(calc, recs)
    years = vals["_years"]
    derived = rce.derive(vals, years, min_cash)
    rev = rce.clean_series(vals["revenue"])
    m = {
        "revenue_last": rev[-1],
        "total_net_income": rce.num(vals.get("total_net_income")),
        "min_cash_balance": rce.num(vals.get("min_cash_balance")),
        "funding_gap": rce.num(vals.get("funding_gap")),
        "peak_funding_need": derived["peak_funding_need"],
        "irr_with_tv": rce.num(vals.get("irr_with_tv")) if not isinstance(vals.get("irr_with_tv"), str) else None,
        "npv": rce.num(vals.get("npv")),
        "dscr_min": rce.num(vals.get("dscr_min")) if not isinstance(vals.get("dscr_min"), str) else None,
        "first_profit_year": derived["first_profit_year"],
        "cumulative_net_income": rce.num(vals.get("total_net_income")),
    }
    return m


def fmt(v, key):
    if v is None:
        return "—"
    if key in ("irr_with_tv",):
        return f"{v * 100:.1f}%"
    if key in ("dscr_min",):
        return f"{v:.2f}"
    if key == "first_profit_year":
        return str(v)
    return f"{v:,.1f}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("assumptions")
    ap.add_argument("-o", "--output", default="sensitivity_outputs.json")
    ap.add_argument("--md", default=None)
    ap.add_argument("--chart", default=None)
    ap.add_argument("--cases", default=None)
    args = ap.parse_args()
    with open(args.assumptions, encoding="utf-8") as f:
        a0 = json.load(f)
    cases = DEFAULT_CASES
    if args.cases:
        with open(args.cases, encoding="utf-8") as f:
            cases = json.load(f)
    soffice = rce.find_soffice()
    if not soffice:
        print("[错误] 未找到 LibreOffice（soffice）", file=sys.stderr)
        sys.exit(2)
    min_cash = float(a0.get("min_cash", 0) or 0)
    tmpdir = tempfile.mkdtemp(prefix="bp_sens_")
    print("[..] 基准")
    base = run_one(copy.deepcopy(a0), 1, soffice, tmpdir, "base", min_cash)
    results = []
    for c in cases:
        print(f"[..] {c['label']}")
        a, sel = apply_case(a0, c)
        m = run_one(a, sel, soffice, tmpdir, c["id"], min_cash)
        delta = {k: (None if (m.get(k) is None or base.get(k) is None or k == "first_profit_year") else m[k] - base[k]) for k in m}
        results.append({"id": c["id"], "label": c["label"], "pair": c.get("pair", c["label"]), "metrics": m, "delta_vs_base": delta})
    shutil.rmtree(tmpdir, ignore_errors=True)
    out = {"unit": a0.get("meta", {}).get("unit", "万元"), "base": base, "cases": results}
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] 输出：{args.output}")

    # Markdown 表
    lines = [f"表 敏感性分析（单因素，{out['unit']}；括号内为相对基准的变化）", "", "| 情形 | " + " | ".join(l for _, l in METRICS) + " |", "|" + "---|" * (len(METRICS) + 1)]
    lines.append("| 基准 | " + " | ".join(fmt(base[k], k) for k, _ in METRICS) + " |")
    for r in results:
        cells = []
        for k, _ in METRICS:
            v = fmt(r["metrics"][k], k)
            d = r["delta_vs_base"].get(k)
            if d is not None and k not in ("first_profit_year",):
                v += f"（{'+' if d >= 0 else ''}{fmt(d, k)}）"
            cells.append(v)
        lines.append(f"| {r['label']} | " + " | ".join(cells) + " |")
    md = "\n".join(lines)
    if args.md:
        with open(args.md, "w", encoding="utf-8") as f:
            f.write(md + "\n")
        print(f"[OK] 表格：{args.md}")
    print(md)

    # 龙卷风图
    if args.chart:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from make_charts import PALETTE, setup_fonts
        setup_fonts("zh")
        pairs = {}
        for r in results:
            pairs.setdefault(r["pair"], []).append(r)
        fig, axes = plt.subplots(1, 2, figsize=(10, 0.6 + 0.5 * len(pairs) + 1))
        for ax, (key, title) in zip(axes, [("npv", f"NPV（{out['unit']}）"), ("min_cash_balance", f"最低期末现金（{out['unit']}）")]):
            rows = []
            for pair, rs in pairs.items():
                ds = [r["delta_vs_base"].get(key) for r in rs if r["delta_vs_base"].get(key) is not None]
                if not ds:
                    continue
                rows.append((pair, min(ds), max(ds)))
            rows.sort(key=lambda t: max(abs(t[1]), abs(t[2])))
            labels = [t[0] for t in rows]
            ax.barh(labels, [t[1] for t in rows], color=PALETTE["orange"], height=0.5, label="不利方向")
            ax.barh(labels, [t[2] for t in rows], color=PALETTE["blue"], height=0.5, label="有利方向")
            ax.axvline(0, color="#999", linewidth=0.8)
            base_v = base.get(key)
            ax.set_title(title + (f"，基准 {base_v:,.0f}" if base_v is not None else ""))
            ax.legend(loc="lower right", fontsize=8)
        fig.text(0.01, 0.01, "来源：财务模型，单因素敏感性", fontsize=7.5, color=PALETTE["muted"])
        fig.tight_layout(rect=(0, 0.03, 1, 1))
        os.makedirs(os.path.dirname(os.path.abspath(args.chart)), exist_ok=True)
        fig.savefig(args.chart, dpi=200, bbox_inches="tight", facecolor="white")
        print(f"[OK] 图：{args.chart}")


if __name__ == "__main__":
    main()
