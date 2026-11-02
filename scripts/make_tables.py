#!/usr/bin/env python3
"""
make_tables.py — 从 model_outputs.json（及可选的 sensitivity_outputs.json）生成正文可直接粘贴的 Markdown 表格，
避免手工转录数字。中文默认万元；英文可换算为 RMB million。

用法:
    python make_tables.py work/model_outputs.json -o work/tables.md [--sensitivity work/sensitivity_outputs.json] [--lang zh|en] [--unit 万元|million]

生成的表（每张带表题，正文按需复制）:
    T1 关键财务指标摘要          T2 利润表            T3 现金流量表         T4 资产负债表
    T5 偿债能力（借款、本息、DSCR） T6 三情景对比        T7 资金来源与总投资（可研口径）
    T8 敏感性分析（有 --sensitivity 时）  T9 分产品线收入  T10 第 N 年月度现金（有月度表时）
"""
import argparse
import json

L = {
    "zh": {
        "t1": "关键财务指标（基准情景，{u}）", "t2": "利润表（基准情景，{u}）", "t3": "现金流量表（基准情景，{u}）", "t4": "资产负债表（年末，基准情景，{u}）",
        "t5": "借款与偿债能力（基准情景，{u}）", "t6": "三情景对比（{u}）", "t7": "资金来源与总投资（可研口径，{u}）", "t8": "敏感性分析（单因素，{u}）",
        "t9": "分产品线收入（基准情景，{u}）", "t10": "第 {k} 年（{y}）月度现金（参考口径，{u}）",
        "item": "项目", "year": "年份", "month": "月份",
        "revenue": "营业收入", "revenue_growth": "收入增长率", "cogs": "营业成本", "gross_profit": "毛利", "gross_margin": "毛利率",
        "opex_销售费用": "销售费用", "opex_管理费用": "管理费用", "opex_研发费用": "研发费用", "opex_其他费用": "其他费用", "ebitda": "EBITDA", "ebitda_margin": "EBITDA 率",
        "depreciation": "折旧与摊销", "ebit": "EBIT", "interest": "利息费用", "ebt": "利润总额", "tax": "所得税", "net_income": "净利润", "net_margin": "净利率",
        "cfo": "经营活动现金流净额", "cfi": "投资活动现金流净额", "cff": "筹资活动现金流净额", "net_cash_change": "现金净增加额", "opening_cash": "期初现金", "closing_cash": "期末现金",
        "unlevered_fcf": "无杠杆自由现金流", "cash": "货币资金", "accounts_receivable": "应收账款", "inventory": "存货", "other_current_assets": "其他流动资产", "current_assets": "流动资产合计",
        "fixed_assets_net": "固定资产净值", "construction_in_progress": "在建工程", "total_assets": "资产总计", "accounts_payable": "应付账款", "debt_balance": "借款余额", "total_liabilities": "负债合计",
        "paid_in_capital": "实收资本", "capital_reserve": "资本公积", "retained_earnings": "未分配利润", "total_equity": "所有者权益合计", "total_liabilities_equity": "负债和所有者权益总计",
        "debt_ratio": "资产负债率", "current_ratio": "流动比率", "roe": "净资产收益率", "equity_in": "股权投入", "grant_in": "政府补助", "loan_drawdown": "借款提款", "loan_repayment": "偿还本金",
        "debt_service": "本息合计", "dscr": "偿债覆盖率 DSCR", "interest_coverage": "利息保障倍数", "capex": "资本开支", "headcount_total": "人员总数（人）",
        "base": "基准", "downside": "下行", "upside": "上行", "metric": "指标",
        "s_rev_last": "末年营业收入", "s_cagr": "收入 CAGR", "s_gm_last": "末年毛利率", "s_ni_last": "末年净利润", "s_cum_ni": "累计净利润", "s_first_profit": "首个盈利年",
        "s_min_cash": "最低期末现金", "s_gap": "资金缺口", "s_peak": "峰值资金需求（无融资口径）", "s_dscr": "经营期最低 DSCR", "s_irr": "项目 IRR（含期末价值）", "s_irr_notv": "项目 IRR（不含期末价值）", "s_npv": "NPV", "s_payback": "回收期（年）",
        "f_capex": "建设投资（资本开支）", "f_int": "建设期利息（未资本化，列示）", "f_wc": "铺底流动资金（首个经营年营运资金占用）", "f_total": "可研口径总投资", "f_equity": "自有资金 / 股权", "f_grant": "政府补助", "f_loan": "银行借款", "f_sources": "资金来源合计", "f_diff": "来源 − 总投资（差额为营运资金储备 / 建设期亏损覆盖）",
        "case": "情形", "note_sens": "括号内为相对基准的变化", "closing_cash_m": "月末现金", "net_cash_m": "净现金流",
    },
    "en": {
        "t1": "Key financial indicators (base case, {u})", "t2": "Income statement (base case, {u})", "t3": "Cash flow statement (base case, {u})", "t4": "Balance sheet (year-end, base case, {u})",
        "t5": "Debt and debt service (base case, {u})", "t6": "Scenario comparison ({u})", "t7": "Funding sources vs total investment (feasibility-study basis, {u})", "t8": "Sensitivity analysis (single factor, {u})",
        "t9": "Revenue by line (base case, {u})", "t10": "Year {k} ({y}) monthly cash (indicative, {u})",
        "item": "Item", "year": "Year", "month": "Month",
        "revenue": "Revenue", "revenue_growth": "Revenue growth", "cogs": "Cost of goods sold", "gross_profit": "Gross profit", "gross_margin": "Gross margin",
        "opex_销售费用": "Selling expenses", "opex_管理费用": "G&A expenses", "opex_研发费用": "R&D expenses", "opex_其他费用": "Other expenses", "ebitda": "EBITDA", "ebitda_margin": "EBITDA margin",
        "depreciation": "Depreciation & amortisation", "ebit": "EBIT", "interest": "Interest expense", "ebt": "Profit before tax", "tax": "Income tax", "net_income": "Net income", "net_margin": "Net margin",
        "cfo": "Operating cash flow", "cfi": "Investing cash flow", "cff": "Financing cash flow", "net_cash_change": "Net change in cash", "opening_cash": "Opening cash", "closing_cash": "Closing cash",
        "unlevered_fcf": "Unlevered free cash flow", "cash": "Cash", "accounts_receivable": "Accounts receivable", "inventory": "Inventory", "other_current_assets": "Other current assets", "current_assets": "Total current assets",
        "fixed_assets_net": "Fixed assets, net", "construction_in_progress": "Construction in progress", "total_assets": "Total assets", "accounts_payable": "Accounts payable", "debt_balance": "Borrowings", "total_liabilities": "Total liabilities",
        "paid_in_capital": "Paid-in capital", "capital_reserve": "Capital reserve", "retained_earnings": "Retained earnings", "total_equity": "Total equity", "total_liabilities_equity": "Total liabilities and equity",
        "debt_ratio": "Debt-to-asset ratio", "current_ratio": "Current ratio", "roe": "Return on equity", "equity_in": "Equity injected", "grant_in": "Government grants", "loan_drawdown": "Loan drawdown", "loan_repayment": "Principal repaid",
        "debt_service": "Debt service", "dscr": "DSCR", "interest_coverage": "Interest coverage", "capex": "Capital expenditure", "headcount_total": "Headcount",
        "base": "Base", "downside": "Downside", "upside": "Upside", "metric": "Metric",
        "s_rev_last": "Final-year revenue", "s_cagr": "Revenue CAGR", "s_gm_last": "Final-year gross margin", "s_ni_last": "Final-year net income", "s_cum_ni": "Cumulative net income", "s_first_profit": "First profitable year",
        "s_min_cash": "Minimum closing cash", "s_gap": "Funding gap", "s_peak": "Peak funding need (before financing)", "s_dscr": "Min. operating DSCR", "s_irr": "Project IRR (with terminal value)", "s_irr_notv": "Project IRR (no terminal value)", "s_npv": "NPV", "s_payback": "Payback (years)",
        "f_capex": "Construction investment (capex)", "f_int": "Construction-period interest (expensed, shown for reference)", "f_wc": "Initial working capital (first operating year)", "f_total": "Total investment, feasibility basis", "f_equity": "Equity / own funds", "f_grant": "Government grants", "f_loan": "Bank loans", "f_sources": "Total funding sources", "f_diff": "Sources − total investment (working-capital buffer / construction losses)",
        "case": "Case", "note_sens": "change vs base in brackets", "closing_cash_m": "Month-end cash", "net_cash_m": "Net cash flow",
    },
}
PCT_KEYS = {"revenue_growth", "gross_margin", "ebitda_margin", "net_margin", "debt_ratio", "roe"}
RATIO_KEYS = {"current_ratio", "dscr", "interest_coverage"}
INT_KEYS = {"headcount_total"}


class Fmt:
    def __init__(self, unit_in, unit_out, lang):
        self.scale = 1.0
        self.lang = lang
        self.unit_label = unit_in
        if unit_in == "万元" and unit_out == "million":
            self.scale = 0.01
            self.unit_label = "RMB million"
        elif unit_in == "万元" and unit_out == "亿元":
            self.scale = 0.0001
            self.unit_label = "亿元"
        elif lang == "en" and unit_in == "万元":
            self.unit_label = "RMB 10k"

    def money(self, v, d=None):
        if v is None:
            return "—"
        x = v * self.scale
        if d is None:
            d = 2 if self.scale != 1.0 else 1
        return f"{x:,.{d}f}"

    @staticmethod
    def pct(v):
        return "—" if v is None else f"{v * 100:.1f}%"

    @staticmethod
    def ratio(v):
        return "—" if v is None else f"{v:.2f}"

    def any(self, key, v):
        if key in PCT_KEYS:
            return self.pct(v)
        if key in RATIO_KEYS:
            return self.ratio(v)
        if key in INT_KEYS:
            return "—" if v is None else f"{v:,.0f}"
        return self.money(v)


def table(title, header, rows):
    out = [title, "", "| " + " | ".join(header) + " |", "|" + "---|" * len(header)]
    for r in rows:
        out.append("| " + " | ".join(str(c) for c in r) + " |")
    return "\n".join(out) + "\n"


def series_table(T, F, title_key, keys, s, years, bold=()):
    rows = []
    for k in keys:
        if k not in s:
            continue
        name = T.get(k, k)
        if k in bold:
            name = f"**{name}**"
        rows.append([name] + [F.any(k, v) for v in s[k]])
    return table(T[title_key].format(u=F.unit_label), [T["item"]] + [str(y) for y in years], rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("outputs")
    ap.add_argument("-o", "--output", default="tables.md")
    ap.add_argument("--sensitivity", default=None)
    ap.add_argument("--lang", default="zh", choices=["zh", "en"])
    ap.add_argument("--unit", default=None, help="输出单位：万元（默认）/ 亿元 / million（RMB million，英文版常用）")
    args = ap.parse_args()
    with open(args.outputs, encoding="utf-8") as f:
        o = json.load(f)
    T = L[args.lang]
    unit_in = o.get("unit", "万元")
    unit_out = args.unit or ("million" if args.lang == "en" else unit_in)
    F = Fmt(unit_in, unit_out, args.lang)
    years = o["years"]
    b = o["scenarios"]["base"]
    s, sm, d = b["series"], b["summary"], b["derived"]
    parts = []

    # T1 summary
    irr = sm.get("irr_with_tv")
    rows = [
        [T["s_rev_last"], F.money(s["revenue"][-1])], [T["s_cagr"] + ("" if sm.get("revenue_cagr") is not None else (f"（自 {d.get('first_operating_year')} 年起）" if args.lang == "zh" else f" (from {d.get('first_operating_year')})")), F.pct(sm.get("revenue_cagr") if sm.get("revenue_cagr") is not None else d.get("revenue_cagr_operating"))], [T["s_gm_last"], F.pct(s["gross_margin"][-1])],
        [T["s_ni_last"], F.money(s["net_income"][-1])], [T["s_cum_ni"], F.money(sm.get("total_net_income"))], [T["s_first_profit"], d.get("first_profit_year") or "—"],
        [T["s_min_cash"], F.money(sm.get("min_cash_balance"))], [T["s_gap"], F.money(sm.get("funding_gap"))], [T["s_peak"], F.money(d.get("peak_funding_need"))],
        [T["s_dscr"], F.ratio(d.get("dscr_min_operating"))], [T["s_irr"], F.pct(irr) if irr is not None else "—"], [T["s_irr_notv"], F.pct(sm.get("irr_no_tv")) if sm.get("irr_no_tv") is not None else "—"],
        [T["s_npv"], F.money(sm.get("npv"))], [T["s_payback"], d.get("payback_years_ufcf") if d.get("payback_years_ufcf") is not None else ("—" if args.lang == "en" else "超出预测期")],
    ]
    parts.append(table(T["t1"].format(u=F.unit_label), [T["metric"], T["base"]], rows))

    parts.append(series_table(T, F, "t2", ["revenue", "revenue_growth", "cogs", "gross_profit", "gross_margin", "opex_销售费用", "opex_管理费用", "opex_研发费用", "opex_其他费用", "ebitda", "ebitda_margin", "depreciation", "ebit", "interest", "ebt", "tax", "net_income", "net_margin"], s, years, bold=("gross_profit", "ebitda", "net_income")))
    parts.append(series_table(T, F, "t3", ["cfo", "cfi", "cff", "net_cash_change", "opening_cash", "closing_cash", "unlevered_fcf"], s, years, bold=("closing_cash",)))
    parts.append(series_table(T, F, "t4", ["cash", "accounts_receivable", "inventory", "other_current_assets", "current_assets", "fixed_assets_net", "construction_in_progress", "total_assets", "accounts_payable", "debt_balance", "total_liabilities", "paid_in_capital", "capital_reserve", "retained_earnings", "total_equity", "total_liabilities_equity", "debt_ratio", "current_ratio"], s, years, bold=("total_assets", "total_liabilities_equity")))
    parts.append(series_table(T, F, "t5", ["loan_drawdown", "loan_repayment", "interest", "debt_service", "debt_balance", "ebitda", "cfo", "dscr", "interest_coverage"], s, years, bold=("dscr",)))

    # T6 scenarios
    codes = [c for c in ("base", "downside", "upside") if c in o["scenarios"]]
    def g(code, path):
        cur = o["scenarios"][code]
        for p in path:
            cur = cur.get(p) if isinstance(cur, dict) else (cur[p] if isinstance(cur, list) else None)
            if cur is None:
                return None
        return cur
    rows = [
        [T["s_rev_last"]] + [F.money(g(c, ["series", "revenue", -1])) for c in codes],
        [T["s_cum_ni"]] + [F.money(g(c, ["summary", "total_net_income"])) for c in codes],
        [T["s_first_profit"]] + [g(c, ["derived", "first_profit_year"]) or "—" for c in codes],
        [T["s_min_cash"]] + [F.money(g(c, ["summary", "min_cash_balance"])) for c in codes],
        [T["s_gap"]] + [F.money(g(c, ["summary", "funding_gap"])) for c in codes],
        [T["s_dscr"]] + [F.ratio(g(c, ["derived", "dscr_min_operating"])) for c in codes],
        [T["s_irr"]] + [(F.pct(g(c, ["summary", "irr_with_tv"])) if g(c, ["summary", "irr_with_tv"]) is not None else "—") for c in codes],
        [T["s_npv"]] + [F.money(g(c, ["summary", "npv"])) for c in codes],
    ]
    parts.append(table(T["t6"].format(u=F.unit_label), [T["metric"]] + [T[c] for c in codes], rows))

    # T7 funding vs total investment
    fs = d.get("funding_sources") or {}
    rows = [
        [T["f_capex"], F.money(d.get("total_capex"))], [T["f_int"], F.money(d.get("construction_period_interest"))], [T["f_wc"], F.money(d.get("initial_working_capital"))],
        [f"**{T['f_total']}**", f"**{F.money(d.get('total_investment_feasibility'))}**"],
        [T["f_equity"], F.money(fs.get("equity"))], [T["f_grant"], F.money(fs.get("grants"))], [T["f_loan"], F.money(fs.get("loans"))],
        [f"**{T['f_sources']}**", f"**{F.money(d.get('funding_sources_total'))}**"],
        [T["f_diff"], F.money((d.get("funding_sources_total") or 0) - (d.get("total_investment_feasibility") or 0))],
    ]
    parts.append(table(T["t7"].format(u=F.unit_label), [T["item"], F.unit_label], rows))

    # T9 revenue by line
    line_keys = [k for k in s if k.startswith("rev_") and not k.startswith("rev_volume_")]
    if line_keys:
        rows = [[k[4:]] + [F.money(v) for v in s[k]] for k in line_keys]
        rows.append([f"**{T['revenue']}**"] + [f"**{F.money(v)}**" for v in s["revenue"]])
        parts.append(table(T["t9"].format(u=F.unit_label), [T["item"]] + [str(y) for y in years], rows))

    # T8 sensitivity
    if args.sensitivity:
        with open(args.sensitivity, encoding="utf-8") as f:
            sens = json.load(f)
        keys = [("revenue_last", T["s_rev_last"]), ("total_net_income", T["s_cum_ni"]), ("min_cash_balance", T["s_min_cash"]), ("funding_gap", T["s_gap"]), ("npv", T["s_npv"]), ("irr_with_tv", T["s_irr"]), ("dscr_min", T["s_dscr"])]
        def fm(k, v):
            if v is None:
                return "—"
            if k == "irr_with_tv":
                return F.pct(v)
            if k == "dscr_min":
                return F.ratio(v)
            return F.money(v)
        rows = [[T["base"]] + [fm(k, sens["base"].get(k)) for k, _ in keys]]
        for c in sens["cases"]:
            cells = []
            for k, _ in keys:
                v = fm(k, c["metrics"].get(k))
                dv = c["delta_vs_base"].get(k)
                if dv is not None and k not in ("dscr_min",):
                    sign = "+" if dv >= 0 else "−"
                    cells.append(f"{v}（{sign}{fm(k, abs(dv))}）" if args.lang == "zh" else f"{v} ({sign}{fm(k, abs(dv))})")
                else:
                    cells.append(v)
            rows.append([c["label"]] + cells)
        parts.append(table(T["t8"].format(u=F.unit_label) + f"，{T['note_sens']}" if args.lang == "zh" else T["t8"].format(u=F.unit_label) + f"; {T['note_sens']}", [T["case"]] + [lab for _, lab in keys], rows))

    # T10 monthly
    m = b.get("monthly_year1")
    if m:
        months = list(range(1, len(m["closing_cash"]) + 1))
        rows = [[T["net_cash_m"]] + [F.money(v) for v in m.get("net_cash", [])], [T["closing_cash_m"]] + [F.money(v) for v in m["closing_cash"]]]
        parts.append(table(T["t10"].format(k=m.get("year_index", 1), y=m.get("year", years[0]), u=F.unit_label), [T["item"]] + [str(x) for x in months], rows))

    text = "\n".join(parts)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"[OK] {args.output}（{len(parts)} 张表，单位 {F.unit_label}）")


if __name__ == "__main__":
    main()
