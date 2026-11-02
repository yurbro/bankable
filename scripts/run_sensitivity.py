#!/usr/bin/env python3
"""
run_sensitivity.py — 单因素敏感性分析：逐个改动关键假设，重算模型，汇总对关键结论的影响。

用法:
    python run_sensitivity.py assumptions.json -o work/sensitivity_outputs.json [--md work/sensitivity.md] [--chart work/charts/sensitivity.png]
        [--cases my_cases.json] [--variants 方案B=alt/assumptions_B.json 方案C=alt/assumptions_C.json] [--lang zh|en]
    --variants 用整套备选假设文件跑“方案比较”，结果以附加情形列入同一张表（内部版第 15 章）。

默认 12 个案例：销量 ±20%、价格 ±10%、单位成本 ±10%、费用 ±10%、资本开支 ±20%、应收账期 ±30 天。
自定义案例文件为数组，元素形如:
    {"id": "price-15", "label": "价格 −15%", "driver": "price", "mult": 0.85}          # driver ∈ volume/price/unit_cost/opex
    {"id": "capex+30", "label": "资本开支 +30%", "patch": "capex", "factor": 1.3}
    {"id": "ar+60",  "label": "应收账期 +60 天", "patch": "ar_days", "delta": 60}
    {"id": "tax15",  "label": "所得税 15%", "patch": "json", "path": "tax.income_tax_rate", "value": 0.15}
    {"id": "rate+1", "label": "贷款利率 +100bp", "patch": "json", "path": "financing.loans.0.rate", "value": 0.055}   # 支持列表索引
    {"id": "nograce","label": "无宽限期", "patch": "json", "path": "financing.loans.0.grace_years", "value": 0}
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
METRICS_ZH = [
    ("revenue_last", "末年营业收入"), ("total_net_income", "累计净利润"), ("min_cash_balance", "最低期末现金"),
    ("funding_gap", "资金缺口"), ("peak_funding_need", "峰值资金需求"), ("irr_with_tv", "项目 IRR（含期末价值）"),
    ("equity_irr", "股权 IRR"), ("npv", "NPV"), ("dscr_min", "经营期最低 DSCR"), ("first_profit_year", "首个盈利年"),
]
METRICS_EN = [
    ("revenue_last", "Final-year revenue"), ("total_net_income", "Cumulative net income"), ("min_cash_balance", "Minimum closing cash"),
    ("funding_gap", "Funding gap"), ("peak_funding_need", "Peak funding need"), ("irr_with_tv", "Project IRR (with TV)"),
    ("equity_irr", "Equity IRR"), ("npv", "NPV"), ("dscr_min", "Min. operating DSCR"), ("first_profit_year", "First profitable year"),
]
METRICS = METRICS_ZH
LABELS = {
    "zh": {"table_title": "表 敏感性分析（单因素，{unit}；括号内为相对基准的变化）", "case": "情形", "base": "基准", "npv": "NPV（{unit}），基准 {base}", "cash": "最低期末现金（{unit}），基准 {base}", "adverse": "不利方向", "favorable": "有利方向", "source": "来源：财务模型，单因素敏感性"},
    "en": {"table_title": "Table: Sensitivity analysis (single factor, {unit}; change vs base in brackets)", "case": "Case", "base": "Base", "npv": "NPV ({unit}), base {base}", "cash": "Minimum closing cash ({unit}), base {base}", "adverse": "Adverse", "favorable": "Favourable", "source": "Source: financial model, single-factor sensitivity"},
}


def _key(cur, k):
    return int(k) if isinstance(cur, list) else k


def set_path(d, path, value):
    """按点路径赋值，支持列表索引：financing.loans.0.rate、capex.1.amount。"""
    keys = path.split(".")
    cur = d
    for k in keys[:-1]:
        cur = cur[_key(cur, k)]
    cur[_key(cur, keys[-1])] = value


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
        "irr_with_tv": (lambda v: None if (v is None or v > 3.0) else v)(rce.num(vals.get("irr_with_tv")) if not isinstance(vals.get("irr_with_tv"), str) else None),
        "npv": rce.num(vals.get("npv")),
        "equity_irr": (lambda v: None if (v is None or v > 3.0) else v)(rce.num(vals.get("equity_irr")) if not isinstance(vals.get("equity_irr"), str) else None),
        "dscr_min": derived.get("dscr_min_operating"),
        "first_profit_year": derived["first_profit_year"],
        "cumulative_net_income": rce.num(vals.get("total_net_income")),
    }
    if m["irr_with_tv"] is not None:
        ufcf = rce.clean_series(vals.get("unlevered_fcf", []))
        first_nonzero = next((x for x in ufcf if x is not None and abs(x) > 1e-9), None)
        if first_nonzero is not None and first_nonzero > 0:
            m["irr_with_tv"] = None  # 无前期净投入，IRR 无意义（与 recalc 保持一致）
    return m


def fmt(v, key):
    if v is None:
        return "—"
    if key in ("irr_with_tv", "equity_irr"):
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
    ap.add_argument("--variants", nargs="*", default=[], help="整套备选假设文件：label=path.json（如 方案B=alt_B/assumptions_B.json），结果作为附加情形列入对比")
    ap.add_argument("--lang", default="zh", choices=["zh", "en"], help="表格与图的语言")
    args = ap.parse_args()
    global METRICS
    METRICS = METRICS_EN if args.lang == "en" else METRICS_ZH
    L = LABELS[args.lang]
    with open(args.assumptions, encoding="utf-8") as f:
        a0 = json.load(f)
    cases = DEFAULT_CASES
    if args.lang == "en":
        en_labels = {"volume-20": ("Volume −20%", "Volume ±20%"), "volume+20": ("Volume +20%", "Volume ±20%"), "price-10": ("Price −10%", "Price ±10%"), "price+10": ("Price +10%", "Price ±10%"),
                     "unit_cost+10": ("Unit cost +10%", "Unit cost ±10%"), "unit_cost-10": ("Unit cost −10%", "Unit cost ±10%"), "opex+10": ("Opex +10%", "Opex ±10%"), "opex-10": ("Opex −10%", "Opex ±10%"),
                     "capex+20": ("Capex +20%", "Capex ±20%"), "capex-20": ("Capex −20%", "Capex ±20%"), "ar+30": ("Receivable days +30", "Receivable days ±30"), "ar-30": ("Receivable days −30", "Receivable days ±30")}
        cases = [dict(c, label=en_labels.get(c["id"], (c["label"], c.get("pair")))[0], pair=en_labels.get(c["id"], (c["label"], c.get("pair")))[1]) for c in DEFAULT_CASES]
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
    for k, spec in enumerate(args.variants):
        label, _, path = spec.partition("=")
        if not path:
            label, path = f"variant{k + 1}", label
        print(f"[..] 备选方案 {label}")
        with open(path, encoding="utf-8") as f:
            av = json.load(f)
        m = run_one(av, 1, soffice, tmpdir, f"variant_{k}", float(av.get("min_cash", min_cash) or 0))
        delta = {kk: (None if (m.get(kk) is None or base.get(kk) is None or kk == "first_profit_year") else m[kk] - base[kk]) for kk in m}
        results.append({"id": f"variant-{label}", "label": label, "pair": label, "variant": True, "metrics": m, "delta_vs_base": delta})
    shutil.rmtree(tmpdir, ignore_errors=True)
    out = {"unit": a0.get("meta", {}).get("unit", "万元"), "base": base, "cases": results}
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] 输出：{args.output}")

    # Markdown 表
    unit_label = out["unit"] if args.lang == "zh" else {"万元": "RMB 10k", "亿元": "RMB 100m"}.get(out["unit"], out["unit"])
    lines = [L["table_title"].format(unit=unit_label), "", f"| {L['case']} | " + " | ".join(l for _, l in METRICS) + " |", "|" + "---|" * (len(METRICS) + 1)]
    lines.append(f"| {L['base']} | " + " | ".join(fmt(base[k], k) for k, _ in METRICS) + " |")
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
        setup_fonts(args.lang)
        pairs = {}
        for r in results:
            if r.get("variant"):
                continue
            pairs.setdefault(r["pair"], []).append(r)
        fig, axes = plt.subplots(1, 2, figsize=(10, 0.6 + 0.5 * len(pairs) + 1))
        for ax, (key, title) in zip(axes, [("npv", L["npv"]), ("min_cash_balance", L["cash"])]):
            rows = []
            for pair, rs in pairs.items():
                ds = [r["delta_vs_base"].get(key) for r in rs if r["delta_vs_base"].get(key) is not None]
                if not ds:
                    continue
                rows.append((pair, min(ds), max(ds)))
            rows.sort(key=lambda t: max(abs(t[1]), abs(t[2])))
            labels = [t[0] for t in rows]
            ax.barh(labels, [t[1] for t in rows], color=PALETTE["orange"], height=0.5, label=L["adverse"])
            ax.barh(labels, [t[2] for t in rows], color=PALETTE["blue"], height=0.5, label=L["favorable"])
            ax.axvline(0, color="#999", linewidth=0.8)
            base_v = base.get(key)
            ax.set_title(title.format(unit=unit_label, base=(f"{base_v:,.0f}" if base_v is not None else "—")))
            ax.legend(loc="lower right", fontsize=8)
        fig.text(0.01, 0.01, L["source"], fontsize=7.5, color=PALETTE["muted"])
        fig.tight_layout(rect=(0, 0.03, 1, 1))
        os.makedirs(os.path.dirname(os.path.abspath(args.chart)), exist_ok=True)
        fig.savefig(args.chart, dpi=200, bbox_inches="tight", facecolor="white")
        print(f"[OK] 图：{args.chart}")


if __name__ == "__main__":
    main()
