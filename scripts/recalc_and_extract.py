#!/usr/bin/env python3
"""
recalc_and_extract.py — 用 LibreOffice 重算财务模型，跑基准/下行/上行三情景，导出 model_outputs.json，
并把“情景对比”写回 Excel，最终交付的 Excel 带缓存计算值（打开即显示数字）。

用法:
    python recalc_and_extract.py 模型.xlsx -o model_outputs.json [--assumptions assumptions.json] [--no-scenarios]

要求: 系统有 soffice（LibreOffice）。没有时脚本会用 --no-recalc 模式提示，无法得到计算值。
输出 JSON 的结构见 references/financial-model-guide.md（“模型输出”一节）。
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SCENARIOS = [(1, "base", "基准"), (2, "downside", "下行"), (3, "upside", "上行")]
SERIES_KEYS = [
    "revenue", "cogs", "gross_profit", "gross_margin", "revenue_growth", "opex_total",
    "opex_销售费用", "opex_管理费用", "opex_研发费用", "opex_其他费用", "headcount_total",
    "capex", "depreciation", "fixed_assets_net", "construction_in_progress", "current_portion_debt",
    "equity_in", "grant_in", "loan_drawdown", "loan_repayment", "interest", "debt_balance", "debt_service",
    "ebitda", "ebitda_margin", "ebit", "ebt", "tax", "net_income", "net_margin", "dividends", "cumulative_net_income",
    "cfo", "cfi", "cff", "net_cash_change", "opening_cash", "closing_cash", "cash_vs_min",
    "unlevered_fcf", "cumulative_ufcf", "terminal_value", "equity_cf",
    "cash", "accounts_receivable", "inventory", "other_current_assets", "current_assets", "total_assets",
    "accounts_payable", "total_liabilities", "paid_in_capital", "capital_reserve", "retained_earnings", "total_equity", "total_liabilities_equity",
    "balance_check", "debt_ratio", "current_ratio", "roe", "interest_coverage", "dscr", "cfo_to_debt_service",
    "breakeven_revenue", "breakeven_ratio",
]
SCALAR_KEYS = [
    "irr_with_tv", "irr_no_tv", "npv", "equity_irr", "min_cash_balance", "min_cash_year", "funding_gap",
    "total_equity_in", "total_grant_in", "total_loan_drawdown", "total_capex", "total_net_income", "total_cfo", "dscr_min", "revenue_cagr",
]


def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice", r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def recalc(soffice, src, outdir, timeout=180):
    os.makedirs(outdir, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    cmd = [soffice, "--headless", "--norestore", f"-env:UserInstallation=file://{profile}", "--convert-to", "xlsx", "--outdir", outdir, src]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    out = os.path.join(outdir, os.path.basename(src))
    if not os.path.exists(out):
        raise RuntimeError("LibreOffice 重算失败：" + (r.stderr or r.stdout))
    shutil.rmtree(profile, ignore_errors=True)
    return out


def read_map(wb):
    ws = wb["_map"]
    recs = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            recs[row[0]] = {"sheet": row[1], "row": int(row[2]), "col": int(row[3]), "n": int(row[4])}
    return recs


def read_values(path, recs):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for key, rec in recs.items():
        if rec["sheet"] not in wb.sheetnames:
            continue
        ws = wb[rec["sheet"]]
        vals = [ws.cell(row=rec["row"], column=rec["col"] + i).value for i in range(rec["n"])]
        out[key] = vals[0] if rec["n"] == 1 else vals
    years_ws = wb["利润表"]
    out["_years"] = [years_ws.cell(row=3, column=3 + i).value for i in range(recs["revenue"]["n"])]
    return out


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def clean_series(vals):
    return [num(v) for v in vals] if isinstance(vals, list) else num(vals)


def derive(vals, years, min_cash):
    """Python 侧派生指标：回收期、首个盈利年、峰值资金需求等。"""
    rev = clean_series(vals.get("revenue", []))
    ni = clean_series(vals.get("net_income", []))
    cfo = clean_series(vals.get("cfo", []))
    cfi = clean_series(vals.get("cfi", []))
    ufcf = clean_series(vals.get("unlevered_fcf", []))
    closing = clean_series(vals.get("closing_cash", []))
    opening = num(vals.get("opening_cash", [0])[0] if isinstance(vals.get("opening_cash"), list) else vals.get("opening_cash")) or 0.0
    n = len(years)

    def first_year(series, cond):
        for i, v in enumerate(series):
            if v is not None and cond(v):
                return years[i]
        return None

    # 回收期（无杠杆自由现金流，线性插值）
    payback = None
    cum = 0.0
    prev_cum = 0.0
    for i, v in enumerate(ufcf):
        prev_cum = cum
        cum += v or 0.0
        if cum >= 0 and i > 0 and prev_cum < 0:
            frac = (-prev_cum) / (v if v else 1)
            payback = round(i + frac, 1)  # 以年为单位（第 i 年末为 i+1）
            break
    # 峰值资金需求：不含任何融资时为保持最低现金所需的外部资金
    peak_need = 0.0
    cum_op = opening
    for i in range(n):
        cum_op += (cfo[i] or 0.0) + (cfi[i] or 0.0)
        peak_need = max(peak_need, min_cash - cum_op)
    neg_years = [years[i] for i in range(n) if closing[i] is not None and closing[i] < 0]
    below_min_years = [years[i] for i in range(n) if closing[i] is not None and closing[i] < min_cash]
    dscr = vals.get("dscr", [])
    dscr_vals = [(years[i], v) for i, v in enumerate(dscr) if num(v) is not None]
    dscr_below = [y for y, v in dscr_vals if v < 1.2]
    repay = clean_series(vals.get("loan_repayment", []))
    # 经营期 DSCR：只看有还本的年份（建设期/宽限期 EBITDA 为负时的 DSCR 对读者无意义）
    dscr_oper = [(years[i], round(v, 2)) for i, v in enumerate(dscr) if num(v) is not None and i < len(repay) and (repay[i] or 0) > 0]
    gm_all = clean_series(vals.get("gross_margin", []))
    gm = [g for g, r in zip(gm_all, rev) if g is not None and r]  # 忽略无收入年份
    growth_all = clean_series(vals.get("revenue_growth", []))
    growth = [g for i, g in enumerate(growth_all) if g is not None and i > 0 and rev[i - 1]]
    # 建设期：首个有收入年份之前的年份
    first_rev_idx = next((i for i, r in enumerate(rev) if r), None)
    construction_years = years[:first_rev_idx] if first_rev_idx else []
    interest = clean_series(vals.get("interest", []))
    construction_interest = sum((interest[i] or 0) for i in range(len(construction_years)))
    capex_s = clean_series(vals.get("capex", []))
    total_capex = sum(x or 0 for x in capex_s)
    # 铺底流动资金：首个经营年的营运资金占用（应收 + 存货 + 其他流动资产 − 应付）
    ar = clean_series(vals.get("accounts_receivable", [])); inv = clean_series(vals.get("inventory", []))
    oca = clean_series(vals.get("other_current_assets", [])); ap = clean_series(vals.get("accounts_payable", []))
    fi = first_rev_idx or 0
    initial_wc = ((ar[fi] or 0) + (inv[fi] or 0) + (oca[fi] or 0) - (ap[fi] or 0)) if fi < len(ar) else 0.0
    first_oper_year = years[first_rev_idx] if first_rev_idx is not None else None
    cagr_oper = None
    if first_rev_idx is not None and rev[first_rev_idx] and n - 1 - first_rev_idx >= 1 and rev[-1] and rev[-1] > 0:
        cagr_oper = round((rev[-1] / rev[first_rev_idx]) ** (1 / (n - 1 - first_rev_idx)) - 1, 4)
    eq_in = sum(x or 0 for x in clean_series(vals.get("equity_in", []))); gr_in = sum(x or 0 for x in clean_series(vals.get("grant_in", []))); ln_in = sum(x or 0 for x in clean_series(vals.get("loan_drawdown", [])))
    return {
        "first_operating_year": first_oper_year,
        "revenue_cagr_operating": cagr_oper,
        "dscr_operating_by_year": dscr_oper,
        "dscr_min_operating": min((v for _, v in dscr_oper), default=None),
        "construction_years": construction_years,
        "construction_period_interest": round(construction_interest, 1),
        "total_capex": round(total_capex, 1),
        "initial_working_capital": round(initial_wc, 1),
        "total_investment_feasibility": round(total_capex + construction_interest + max(initial_wc, 0), 1),
        "funding_sources_total": round(eq_in + gr_in + ln_in, 1),
        "funding_sources": {"equity": round(eq_in, 1), "grants": round(gr_in, 1), "loans": round(ln_in, 1)},
        "first_profit_year": first_year(ni, lambda v: v > 0),
        "first_positive_cfo_year": first_year(cfo, lambda v: v > 0),
        "payback_years_ufcf": payback,
        "peak_funding_need": round(peak_need, 1),
        "negative_cash_years": neg_years,
        "cash_below_min_years": below_min_years,
        "dscr_by_year": [[y, round(v, 2)] for y, v in dscr_vals],
        "dscr_below_1_2_years": dscr_below,
        "gross_margin_range": [round(min(gm), 3), round(max(gm), 3)] if gm else None,
        "max_revenue_growth": round(max(growth), 3) if growth else None,
        "revenue_first_last": [rev[0], rev[-1]] if rev else None,
    }


def fmt_money(v, unit):
    if v is None:
        return None
    if abs(v) >= 100000 and unit == "万元":
        return f"{v / 10000:,.2f} 亿元"
    return f"{v:,.1f} {unit}"


def fmt_pct(v):
    return None if v is None else f"{v * 100:.1f}%"


def build_outputs(results, recs, unit, assumptions, model_path):
    base = results["base"]
    years = base["_years"]
    min_cash = num(base.get("min_cash")) or 0.0
    out = {
        "generated": date.today().isoformat(),
        "model_file": os.path.basename(model_path),
        "project_name": base.get("project_name"),
        "unit": unit,
        "years": years,
        "min_cash_requirement": min_cash,
        "scenarios": {},
        "checks": {},
        "narrative_numbers": {},
    }
    for code, label in [("base", "基准"), ("downside", "下行"), ("upside", "上行")]:
        if code not in results:
            continue
        vals = results[code]
        series = {k: clean_series(vals[k]) for k in SERIES_KEYS if k in vals}
        for k, v in vals.items():  # 其余按年序列（分收入线 rev_*/cogs_*/rev_volume_*、情景乘数等）
            if k in series or k.startswith("_") or k.startswith("monthly") or not isinstance(v, list) or len(v) != len(years):
                continue
            series[k] = clean_series(v)
        scalars = {k: num(vals.get(k)) if not isinstance(vals.get(k), str) else None for k in SCALAR_KEYS if k in vals}
        scalars["min_cash_year"] = vals.get("min_cash_year")
        # IRR 只在“先投入后回收”的现金流上有意义：首个非零无杠杆现金流为正，或结果离谱（>300%），置空并说明
        ufcf = clean_series(vals.get("unlevered_fcf", []))
        first_nonzero = next((x for x in ufcf if x is not None and abs(x) > 1e-9), None)
        irr_note = None
        if first_nonzero is not None and first_nonzero > 0:
            irr_note = "现金流无前期净投入（存量业务或投资被经营现金流覆盖），项目 IRR 无意义；请用有无对比的增量现金流或 NPV/回收期表述"
        for k in ("irr_with_tv", "irr_no_tv", "equity_irr"):
            v = scalars.get(k)
            if v is not None and (irr_note or v > 3.0):
                scalars[k] = None
                irr_note = irr_note or f"{k} 计算值 {v:.0%} 超出合理范围，已置空"
        scalars["irr_note"] = irr_note
        derived = derive(vals, years, min_cash)
        out["scenarios"][code] = {"label": label, "series": series, "summary": scalars, "derived": derived}
        if "monthly_closing_cash" in vals:
            mc = clean_series(vals["monthly_closing_cash"])
            m_idx = int(num(vals.get("monthly_year_index")) or 1)
            annual_y1 = series["closing_cash"][m_idx - 1] if series.get("closing_cash") else None
            diff = (mc[-1] - annual_y1) if (mc[-1] is not None and annual_y1 is not None) else None
            out["scenarios"][code]["monthly_year1"] = {
                "year_index": m_idx,
                "year": years[m_idx - 1],
                "closing_cash": mc,
                "net_cash": clean_series(vals.get("monthly_net_cash", [])),
                "min_closing_cash": min(x for x in mc if x is not None),
                "min_month": mc.index(min(x for x in mc if x is not None)) + 1,
                "collection_lag_months": num(vals.get("monthly_collection_lag")),
                "payment_lag_months": num(vals.get("monthly_payment_lag")),
                "annual_year1_closing_cash": annual_y1,
                "difference_vs_annual": diff,
                "explanation": "月度表按整月收付款滞后分摊且不含存货占用，年度表按周转天数计算年末应收/应付/存货，所以 12 月末现金与年度表期末现金不同；月度表用于判断现金节奏与最低点，正式口径以年度表为准。",
            }
    b = out["scenarios"]["base"]
    bc = clean_series(base.get("balance_check", []))
    max_diff = max(abs(x) for x in bc if x is not None) if bc else None
    out["checks"] = {
        "balance_sheet_balances": (max_diff is not None and max_diff < 0.01),
        "max_balance_diff": max_diff,
        "negative_cash_years_base": b["derived"]["negative_cash_years"],
        "cash_below_min_years_base": b["derived"]["cash_below_min_years"],
        "funding_gap_base": b["summary"].get("funding_gap"),
        "funding_gap_downside": out["scenarios"].get("downside", {}).get("summary", {}).get("funding_gap"),
        "dscr_min_base": b["summary"].get("dscr_min"),
        "dscr_min_operating_base": b["derived"]["dscr_min_operating"],
        "dscr_below_1_2_years_base": b["derived"]["dscr_below_1_2_years"],
        "dscr_min_downside": out["scenarios"].get("downside", {}).get("summary", {}).get("dscr_min"),
        "dscr_min_operating_downside": out["scenarios"].get("downside", {}).get("derived", {}).get("dscr_min_operating"),
        "construction_years": b["derived"]["construction_years"],
        "total_investment_feasibility": b["derived"]["total_investment_feasibility"],
        "funding_sources_total": b["derived"]["funding_sources_total"],
        "gross_margin_range_base": b["derived"]["gross_margin_range"],
        "max_revenue_growth_base": b["derived"]["max_revenue_growth"],
        "warnings": [],
    }
    w = out["checks"]["warnings"]
    if not out["checks"]["balance_sheet_balances"]:
        w.append(f"资产负债表不平（最大差额 {max_diff}），请检查期初余额或模型")
    if b["derived"]["negative_cash_years"]:
        w.append(f"基准情景现金为负：{b['derived']['negative_cash_years']}，必须调整融资或支出安排")
    elif b["derived"]["cash_below_min_years"]:
        w.append(f"基准情景现金低于最低现金要求（{min_cash}）：{b['derived']['cash_below_min_years']}，资金缺口 {b['summary'].get('funding_gap'):,.1f}")
    if out["checks"]["funding_gap_downside"]:
        w.append(f"下行情景存在资金缺口 {out['checks']['funding_gap_downside']:,.1f}，建议在风险章节说明应对（备用授信、分期投入、费用弹性）")
    oper_below = [y for y, v in b["derived"]["dscr_operating_by_year"] if v < 1.2]
    if oper_below:
        w.append(f"基准情景经营期 DSCR 低于 1.2 的年份：{oper_below}（银行通常要求 ≥1.2–1.5）")
    if b["derived"]["construction_years"]:
        w.append(f"建设期 {b['derived']['construction_years']} 无收入：正文用经营期 DSCR（dscr_min_operating）而不是含建设期的最低值；可研口径总投资 = 建设投资 {b['derived']['total_capex']:,.1f} + 建设期利息 {b['derived']['construction_period_interest']:,.1f} + 铺底流动资金 {b['derived']['initial_working_capital']:,.1f} = {b['derived']['total_investment_feasibility']:,.1f}（模型未将建设期利息资本化，正文需说明口径）")
    fs, ti = b["derived"]["funding_sources_total"], b["derived"]["total_investment_feasibility"]
    if ti and fs and abs(fs - ti) / ti > 0.15:
        w.append(f"资金来源合计 {fs:,.1f} 与可研口径总投资 {ti:,.1f} 相差超过 15%，请在融资章节解释差额（营运资金储备、建设期亏损、分期融资等）")
    if assumptions and (assumptions.get("financing", {}) or {}).get("use_of_funds"):
        uof_total = sum(float(u.get("amount", 0) or 0) for u in assumptions["financing"]["use_of_funds"])
        if fs and abs(uof_total - fs) / fs > 0.01:
            w.append(f"use_of_funds 合计 {uof_total:,.1f} ≠ 资金来源合计 {fs:,.1f}（股权 + 补助 + 借款），资金用途表与融资方案不一致")
    gmr = b["derived"]["gross_margin_range"]
    if gmr and (gmr[0] < 0.05 or gmr[1] > 0.85):
        w.append(f"毛利率区间 {gmr} 偏离常见范围，请核对单价/成本假设")
    rev = b["series"].get("revenue") or []
    growth = b["series"].get("revenue_growth") or []
    for i, g in enumerate(growth):
        if g is not None and g > 1.0 and i > 0 and rev[i - 1] and rev[-1]:
            if rev[i - 1] >= 0.15 * rev[-1]:
                w.append(f"基准情景 {years[i]} 年收入增长 {g:.0%}（上年已达末年收入的 {rev[i-1]/rev[-1]:.0%}），需在正文说明产能与获客机制")
            else:
                w.append(f"基准情景 {years[i]} 年收入增长 {g:.0%} 来自小基数（上年收入仅为末年的 {rev[i-1]/rev[-1]:.0%}），正文一句话说明爬坡机制即可")

    # 叙述用的格式化数字
    s = b["series"]
    sm = b["summary"]
    d = b["derived"]
    nn = {
        "years": years,
        "revenue_by_year": [fmt_money(v, unit) for v in s["revenue"]],
        "net_income_by_year": [fmt_money(v, unit) for v in s["net_income"]],
        "ebitda_by_year": [fmt_money(v, unit) for v in s["ebitda"]],
        "closing_cash_by_year": [fmt_money(v, unit) for v in s["closing_cash"]],
        "gross_margin_by_year": [fmt_pct(v) for v in s["gross_margin"]],
        "net_margin_by_year": [fmt_pct(v) for v in s["net_margin"]],
        "revenue_cagr": fmt_pct(sm.get("revenue_cagr")),
        "revenue_cagr_operating": fmt_pct(d.get("revenue_cagr_operating")),
        "revenue_cagr_note": (f"首年无收入，CAGR 以首个经营年 {d.get('first_operating_year')} 起算" if d.get("construction_years") else None),
        "total_capex": fmt_money(sm.get("total_capex"), unit),
        "total_equity_in": fmt_money(sm.get("total_equity_in"), unit),
        "total_loan_drawdown": fmt_money(sm.get("total_loan_drawdown"), unit),
        "total_grant_in": fmt_money(sm.get("total_grant_in"), unit),
        "external_funding_total": fmt_money((sm.get("total_equity_in") or 0) + (sm.get("total_loan_drawdown") or 0) + (sm.get("total_grant_in") or 0), unit),
        "peak_funding_need": fmt_money(d["peak_funding_need"], unit),
        "min_cash_balance": fmt_money(sm.get("min_cash_balance"), unit),
        "min_cash_year": sm.get("min_cash_year"),
        "funding_gap": fmt_money(sm.get("funding_gap"), unit),
        "first_profit_year": d["first_profit_year"],
        "first_positive_cfo_year": d["first_positive_cfo_year"],
        "payback_years_ufcf": d["payback_years_ufcf"],
        "irr_note": sm.get("irr_note"),
        "irr_with_tv": fmt_pct(sm.get("irr_with_tv")),
        "irr_no_tv": fmt_pct(sm.get("irr_no_tv")),
        "equity_irr": fmt_pct(sm.get("equity_irr")),
        "npv": fmt_money(sm.get("npv"), unit),
        "dscr_min": None if sm.get("dscr_min") is None else f"{sm['dscr_min']:.2f}",
        "dscr_min_operating": None if d.get("dscr_min_operating") is None else f"{d['dscr_min_operating']:.2f}",
        "construction_years": d.get("construction_years"),
        "construction_period_interest": fmt_money(d.get("construction_period_interest"), unit),
        "initial_working_capital": fmt_money(d.get("initial_working_capital"), unit),
        "total_investment_feasibility": fmt_money(d.get("total_investment_feasibility"), unit),
        "funding_sources_total": fmt_money(d.get("funding_sources_total"), unit),
        "cumulative_net_income": fmt_money(sm.get("total_net_income"), unit),
        "cumulative_cfo": fmt_money(sm.get("total_cfo"), unit),
    }
    for code in ("downside", "upside"):
        if code in out["scenarios"]:
            sc = out["scenarios"][code]
            nn[f"{code}_revenue_last_year"] = fmt_money(sc["series"]["revenue"][-1], unit)
            nn[f"{code}_cumulative_net_income"] = fmt_money(sc["summary"].get("total_net_income"), unit)
            nn[f"{code}_min_cash_balance"] = fmt_money(sc["summary"].get("min_cash_balance"), unit)
            nn[f"{code}_funding_gap"] = fmt_money(sc["summary"].get("funding_gap"), unit)
            nn[f"{code}_dscr_min"] = None if sc["summary"].get("dscr_min") is None else f"{sc['summary']['dscr_min']:.2f}"
            nn[f"{code}_irr_with_tv"] = fmt_pct(sc["summary"].get("irr_with_tv"))
    out["narrative_numbers"] = nn
    if assumptions:
        out["assumptions_meta"] = assumptions.get("meta")
        out["scenario_labels"] = {k: (assumptions.get("scenarios", {}) or {}).get(k, {}).get("label") for k in ("downside", "upside")}
        out["use_of_funds"] = (assumptions.get("financing", {}) or {}).get("use_of_funds")
    return out


def write_scenario_sheet(wb, out, recs):
    ws = wb["情景对比"]
    # 清空旧内容
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for c in row:
            c.value = None
    unit = out["unit"]
    bold = Font(bold=True)
    hdr_fill = PatternFill("solid", fgColor="D9E1F2")
    ws["A3"] = f"三情景关键结果（静态值，由 recalc_and_extract.py 于 {out['generated']} 写入；修改假设后需重新运行脚本）"
    ws["A3"].font = Font(italic=True, color="666666", size=9)
    r = 5
    headers = ["指标", "基准", "下行", "上行", "单位"]
    for j, h in enumerate(headers):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.font = bold
        c.fill = hdr_fill
    r += 1
    codes = ["base", "downside", "upside"]
    years = out["years"]

    def val(code, path):
        sc = out["scenarios"].get(code)
        if not sc:
            return None
        cur = sc
        for p in path:
            if isinstance(cur, dict):
                cur = cur.get(p)
            elif isinstance(cur, list) and isinstance(p, int):
                cur = cur[p] if -len(cur) <= p < len(cur) else None
            else:
                return None
        return cur

    rows = [
        (f"营业收入（{years[0]}）", ["series", "revenue", 0], unit, "money"),
        (f"营业收入（{years[-1]}）", ["series", "revenue", -1], unit, "money"),
        ("收入 CAGR", ["summary", "revenue_cagr"], "%", "pct"),
        (f"毛利率（{years[-1]}）", ["series", "gross_margin", -1], "%", "pct"),
        (f"EBITDA（{years[-1]}）", ["series", "ebitda", -1], unit, "money"),
        (f"净利润（{years[-1]}）", ["series", "net_income", -1], unit, "money"),
        ("预测期累计净利润", ["summary", "total_net_income"], unit, "money"),
        ("首个盈利年份", ["derived", "first_profit_year"], "年", "raw"),
        ("首个经营现金流为正年份", ["derived", "first_positive_cfo_year"], "年", "raw"),
        ("最低期末现金", ["summary", "min_cash_balance"], unit, "money"),
        ("最低现金出现年份", ["summary", "min_cash_year"], "年", "raw"),
        ("资金缺口（相对最低现金要求）", ["summary", "funding_gap"], unit, "money"),
        ("峰值资金需求（无融资口径）", ["derived", "peak_funding_need"], unit, "money"),
        ("最低 DSCR（全部年份）", ["summary", "dscr_min"], "x", "ratio"),
        ("最低 DSCR（经营期，有还本年份）", ["derived", "dscr_min_operating"], "x", "ratio"),
        ("项目 IRR（含期末价值）", ["summary", "irr_with_tv"], "%", "pct"),
        ("项目 IRR（不含期末价值）", ["summary", "irr_no_tv"], "%", "pct"),
        ("项目 NPV", ["summary", "npv"], unit, "money"),
        ("回收期（无杠杆 FCF，年）", ["derived", "payback_years_ufcf"], "年", "ratio"),
    ]
    fmts = {"money": '#,##0.0;[Red]-#,##0.0;"-"', "pct": "0.0%", "ratio": "0.00", "raw": "0"}
    for label, path, u, kind in rows:
        ws.cell(row=r, column=1, value=label)
        for j, code in enumerate(codes):
            v = val(code, path)
            c = ws.cell(row=r, column=2 + j, value=v)
            c.number_format = fmts[kind]
            c.alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=5, value=u).font = Font(color="666666", size=9)
        r += 1
    r += 1
    for series_key, label in [("revenue", "营业收入"), ("net_income", "净利润"), ("closing_cash", "期末现金"), ("cfo", "经营活动现金流")]:
        c = ws.cell(row=r, column=1, value=f"{label}（逐年，{unit}）")
        c.font = bold
        c.fill = hdr_fill
        for j, y in enumerate(years):
            c = ws.cell(row=r, column=2 + j, value=y)
            c.font = bold
            c.fill = hdr_fill
        r += 1
        for code, lab in zip(codes, ["基准", "下行", "上行"]):
            ws.cell(row=r, column=1, value=lab)
            ser = val(code, ["series", series_key]) or []
            for j, v in enumerate(ser):
                c = ws.cell(row=r, column=2 + j, value=v)
                c.number_format = fmts["money"]
            r += 1
        r += 1
    ws.column_dimensions["A"].width = 34
    for col in "BCDEFGHIJKL":
        ws.column_dimensions[col].width = 14


def main():
    ap = argparse.ArgumentParser(description="重算财务模型、跑三情景并导出 model_outputs.json")
    ap.add_argument("model")
    ap.add_argument("-o", "--output", default="model_outputs.json")
    ap.add_argument("--assumptions", help="可选：assumptions.json，用于把 use_of_funds、情景说明等带入输出")
    ap.add_argument("--no-scenarios", action="store_true", help="只算当前情景")
    args = ap.parse_args()

    soffice = find_soffice()
    if not soffice:
        print("[错误] 未找到 LibreOffice（soffice）。无法重算公式。请安装 LibreOffice，或在 Excel 中打开模型后手动读取数字。", file=sys.stderr)
        sys.exit(2)

    model_path = os.path.abspath(args.model)
    wb_formula = openpyxl.load_workbook(model_path)
    recs = read_map(wb_formula)
    sel = recs["scenario_selector"]
    tmpdir = tempfile.mkdtemp(prefix="bp_recalc_")
    results = {}
    scenarios = SCENARIOS if not args.no_scenarios else SCENARIOS[:1]
    for code_num, code, label in scenarios:
        wb_formula["假设"].cell(row=sel["row"], column=sel["col"]).value = code_num
        tmp_in = os.path.join(tmpdir, f"scenario_{code}", os.path.basename(model_path))
        os.makedirs(os.path.dirname(tmp_in), exist_ok=True)
        wb_formula.save(tmp_in)
        calc = recalc(soffice, tmp_in, os.path.join(tmpdir, f"calc_{code}"))
        results[code] = read_values(calc, recs)
        print(f"[OK] 情景「{label}」重算完成")
    wb_formula["假设"].cell(row=sel["row"], column=sel["col"]).value = 1

    assumptions = None
    if args.assumptions:
        with open(args.assumptions, encoding="utf-8") as f:
            assumptions = json.load(f)
    unit = results["base"].get("unit") or "万元"
    out = build_outputs(results, recs, unit, assumptions, model_path)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    # 写回情景对比并生成带缓存值的最终文件
    write_scenario_sheet(wb_formula, out, recs)
    tmp_final = os.path.join(tmpdir, "final", os.path.basename(model_path))
    os.makedirs(os.path.dirname(tmp_final), exist_ok=True)
    wb_formula.save(tmp_final)
    final = recalc(soffice, tmp_final, os.path.join(tmpdir, "final_calc"))
    shutil.copyfile(final, model_path)
    shutil.rmtree(tmpdir, ignore_errors=True)

    print(f"[OK] 输出：{args.output}")
    print(f"[OK] 模型已更新（含情景对比与缓存值）：{model_path}")
    print("检查结果：")
    def _r(v):
        if isinstance(v, float):
            return round(v, 2)
        if isinstance(v, list):
            return [_r(x) for x in v]
        return v
    for k, v in out["checks"].items():
        if k != "warnings":
            print(f"  - {k}: {_r(v)}")
    for wmsg in out["checks"]["warnings"]:
        print(f"  ! {wmsg}")
    if not out["checks"]["warnings"]:
        print("  （无警告）")


if __name__ == "__main__":
    main()
