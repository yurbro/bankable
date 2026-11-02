#!/usr/bin/env python3
"""
check_consistency.py — 核对正文（plan.md 或 .docx）里的财务数字是否与 model_outputs.json 一致。

用法:
    python check_consistency.py plan.md model_outputs.json [-o consistency_report.md] [--tolerance 0.006]
        [--extra sensitivity_outputs.json ...]

做两件事:
 1. 必提数字：模型的关键结论（首年/末年收入、末年净利润、资金需求、最低现金、IRR、DSCR 等）是否在正文中出现。
 2. 可疑数字：财务/融资相关章节里出现的“数字+万元/亿元/%”，若与模型任何数值（含派生值：差额、占比、累计、情景差异、
    资金用途占比，以及 --extra 传入的敏感性等输出）都对不上（容差默认 0.6%），列为“需复核”。
    紧跟登记册/来源标记（[A-xx] [F-xx] [S-xx] 等）的数字视为已有出处，自动跳过。
    仍是启发式检查：行业数据会被误报，人工确认即可；但不能有模型里没有的“预测数字”。
退出码 0 = 通过（可能有需复核项），1 = 有必提数字缺失。
"""
import argparse
import json
import os
import re
import sys


def load_text(path):
    if path.lower().endswith(".docx"):
        from docx import Document
        doc = Document(path)
        parts = [p.text for p in doc.paragraphs]
        for t in doc.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    with open(path, encoding="utf-8") as f:
        return f.read()


def _find(text, s, needs_yi=False):
    """在正文中找数字写法：前后不能紧邻其他数字/小数点；亿元变体必须紧跟“亿”，million 变体必须紧跟 million。"""
    if needs_yi == "million":
        tail = r"\s*(?:million|m\b|mn\b)"
    elif needs_yi:
        tail = r"\s*亿"
    else:
        tail = r"(?![\d])"
    pat = r"(?<![\d.,])" + re.escape(s) + tail
    return re.search(pat, text, flags=re.I) is not None


def _meaningful(s):
    """过滤掉太短、几乎任何正文都会包含的写法（如 "0"、"1"、"0.0"）。"""
    digits = s.replace(",", "").replace(".", "").replace("-", "").lstrip("0")
    return len(digits) >= 2


def money_variants(v, unit):
    """一个数值可能在正文中出现的写法（万元口径）：返回 {(写法, 是否亿元写法)}。零值不要求提及。"""
    out = set()
    if v is None or abs(v) < 0.05:
        return out
    for d in (0, 1, 2):
        for s in (f"{v:,.{d}f}", f"{v:.{d}f}"):
            if _meaningful(s):
                out.add((s, False))
    if unit == "万元" and abs(v) >= 1000:
        for d in (1, 2, 3):
            out.add((f"{v / 10000:.{d}f}", True))  # 亿元写法，必须紧跟“亿”
    if unit == "万元" and abs(v) >= 10:
        for d in (1, 2):
            out.add((f"{v / 100:,.{d}f}", "million"))  # 英文 RMB xx.x million 写法
    if v < 0:  # 负数去掉符号的写法（“亏损 47.8 万元”）
        out |= money_variants(-v, unit)
    return out


def pct_variants(v):
    out = set()
    if v is None:
        return out
    for d in (0, 1, 2):
        s = f"{v * 100:.{d}f}"
        if _meaningful(s) or d >= 1:
            out.add((s + "%", False))
            out.add((s + " %", False))
    return out


def collect_model_numbers(o):
    """返回 (money_values, pct_values, ratio_values) 的集合，用于可疑数字匹配。"""
    money, pct, ratio = set(), set(), set()
    for code, sc in o["scenarios"].items():
        for k, ser in sc["series"].items():
            for v in (ser if isinstance(ser, list) else [ser]):
                if v is None:
                    continue
                if any(t in k for t in ("margin", "growth", "ratio", "roe", "cagr")):
                    pct.add(v)
                elif k in ("dscr", "interest_coverage", "cfo_to_debt_service", "current_ratio"):
                    ratio.add(v)
                else:
                    money.add(v)
        for k, v in sc["summary"].items():
            if v is None or isinstance(v, str):
                continue
            if k in ("irr_with_tv", "irr_no_tv", "equity_irr", "revenue_cagr"):
                pct.add(v)
            elif k in ("dscr_min",):
                ratio.add(v)
            elif k == "min_cash_year":
                continue
            else:
                money.add(v)
        d = sc.get("derived", {})
        for k in ("peak_funding_need", "construction_period_interest", "total_capex", "initial_working_capital", "total_investment_feasibility", "funding_sources_total"):
            if d.get(k) is not None:
                money.add(d[k])
        for k, v in (d.get("funding_sources") or {}).items():
            money.add(v)
        if d.get("dscr_min_operating") is not None:
            ratio.add(d["dscr_min_operating"])
        for y, v in d.get("dscr_by_year", []) + d.get("dscr_operating_by_year", []):
            ratio.add(v)
        m = sc.get("monthly_year1")
        if m:
            money.update(x for x in m.get("closing_cash", []) if x is not None)
            money.add(m.get("min_closing_cash"))
    if o.get("use_of_funds"):
        amounts = [float(u["amount"]) for u in o["use_of_funds"]]
        total = sum(amounts)
        money.update(amounts)
        money.add(total)
        if total:
            pct.update(x / total for x in amounts)
    if o.get("min_cash_requirement"):
        money.add(float(o["min_cash_requirement"]))
    # 派生值：逐年差额、累计、占收入比、情景间差异
    base = o["scenarios"].get("base", {})
    bs = base.get("series", {})
    rev = bs.get("revenue") or []
    for key in ("revenue", "net_income", "ebitda", "closing_cash", "cfo", "gross_profit", "opex_total", "capex", "debt_balance"):
        ser = bs.get(key) or []
        vals = [v for v in ser if v is not None]
        for i in range(1, len(ser)):
            if ser[i] is not None and ser[i - 1] is not None:
                money.add(ser[i] - ser[i - 1])
                if ser[i - 1]:
                    pct.add(ser[i] / ser[i - 1] - 1)
        cum = 0.0
        for v in vals:
            cum += v
            money.add(cum)
        money.add(sum(vals))
        if rev and key != "revenue":
            for v, r in zip(ser, rev):
                if v is not None and r:
                    pct.add(v / r)
    for key in ("opex_销售费用", "opex_管理费用", "opex_研发费用", "opex_其他费用", "cogs", "depreciation", "interest", "tax"):
        ser = bs.get(key) or []
        money.add(sum(v for v in ser if v is not None))
        for v, r in zip(ser, rev):
            if v is not None and r:
                pct.add(v / r)
    for code in ("downside", "upside"):
        sc = o["scenarios"].get(code)
        if not sc:
            continue
        for key in ("revenue", "net_income", "closing_cash", "cfo"):
            a, b = bs.get(key) or [], sc["series"].get(key) or []
            for x, y in zip(a, b):
                if x is not None and y is not None:
                    money.add(y - x)
                    if x:
                        pct.add(y / x - 1)
        for key in ("total_net_income", "min_cash_balance", "npv", "funding_gap"):
            x, y = base.get("summary", {}).get(key), sc.get("summary", {}).get(key)
            if isinstance(x, (int, float)) and isinstance(y, (int, float)):
                money.add(y - x)
                if x:
                    pct.add(y / x - 1)
    # 常见派生：合计类
    b = o["scenarios"]["base"]["summary"]
    money.add((b.get("total_equity_in") or 0) + (b.get("total_loan_drawdown") or 0))
    money.add((b.get("total_equity_in") or 0) + (b.get("total_loan_drawdown") or 0) + (b.get("total_grant_in") or 0))
    return money, pct, ratio


def find_in_text(text, variants):
    # 先匹配万元/百分比写法，再匹配亿元写法
    for needs in (False, True, "million"):
        for item in sorted(variants, key=lambda t: -len(t[0])):
            s, yi = item if isinstance(item, tuple) else (item, False)
            if yi != needs or not s:
                continue
            if _find(text, s, needs_yi=yi):
                return s + ("亿" if yi is True else (" million" if yi == "million" else ""))
    return None


NUM_TOKEN_RE = re.compile(r"(-?\d[\d,]*\.?\d*)\s*(亿元|万元|亿|万|%|％|倍|x(?![a-z])|million|billion)", re.I)
SOURCED_RE = re.compile(r"\[(?:[FAETS]-?\d+|S\d+)")


def collect_extra_numbers(path, money, pct, ratio):
    """把任意 JSON（如 sensitivity_outputs.json）里的数值叶子加入匹配集合。"""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    def classify(key, v):
        k = key.lower()
        if isinstance(v, int) and 1990 <= v <= 2100 and ("year" in k or "年" in key):
            return  # 年份本身不是数字证据
        if any(t in k for t in ("irr", "margin", "growth", "cagr", "pct", "percent", "share", "roe", "rate")) and "ratio" not in k:
            pct.add(float(v))
        elif any(t in k for t in ("dscr", "coverage", "ratio", "to_ds", "times", "multiple", "_x", "倍")):
            ratio.add(float(v))
        else:
            money.add(float(v))

    def walk(node, key=""):
        if isinstance(node, dict):
            for k, v in node.items():
                walk(v, k)
        elif isinstance(node, list):
            for v in node:
                walk(v, key)
        elif isinstance(node, (int, float)) and not isinstance(node, bool):
            classify(key, node)
    walk(data)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("doc")
    ap.add_argument("outputs")
    ap.add_argument("-o", "--report", default="consistency_report.md")
    ap.add_argument("--tolerance", type=float, default=0.006)
    ap.add_argument("--extra", nargs="*", default=[], help="其他数值来源 JSON（如 sensitivity_outputs.json）")
    args = ap.parse_args()

    text = load_text(args.doc)
    with open(args.outputs, encoding="utf-8") as f:
        o = json.load(f)
    unit = o.get("unit", "万元")
    base = o["scenarios"]["base"]
    s, sm, d = base["series"], base["summary"], base["derived"]
    years = o["years"]

    first_i = next((i for i, v in enumerate(s["revenue"]) if v), 0)  # 建设期无收入时取首个经营年
    required = [
        (f"首个经营年（{years[first_i]}）营业收入", money_variants(s["revenue"][first_i], unit)),
        (f"末年（{years[-1]}）营业收入", money_variants(s["revenue"][-1], unit)),
        (f"末年（{years[-1]}）净利润", money_variants(s["net_income"][-1], unit)),
        ("股权投入合计 / 借款合计 / 外部资金合计（至少一项）", money_variants(sm.get("total_equity_in"), unit) | money_variants(sm.get("total_loan_drawdown"), unit) | money_variants((sm.get("total_equity_in") or 0) + (sm.get("total_loan_drawdown") or 0), unit)),
        ("最低期末现金或资金缺口或峰值资金需求（至少一项）", money_variants(sm.get("min_cash_balance"), unit) | money_variants(sm.get("funding_gap"), unit) | money_variants(d.get("peak_funding_need"), unit)),
        (f"末年毛利率", pct_variants(s["gross_margin"][-1])),
    ]
    if sm.get("irr_with_tv") is not None:
        required.append(("项目 IRR（含期末价值）", pct_variants(sm["irr_with_tv"])))
    if d.get("dscr_min_operating") is not None:
        required.append(("经营期最低 DSCR（两位小数）", {(f"{d['dscr_min_operating']:.2f}", False)}))
    elif sm.get("dscr_min") is not None and d.get("dscr_by_year"):
        required.append(("最低 DSCR（两位小数）", {(f"{sm['dscr_min']:.2f}", False)}))
    if "downside" in o["scenarios"]:
        ds = o["scenarios"]["downside"]
        required.append(("下行情景末年收入或累计净利润或最低现金（至少一项）", money_variants(ds["series"]["revenue"][-1], unit) | money_variants(ds["summary"].get("total_net_income"), unit) | money_variants(ds["summary"].get("min_cash_balance"), unit) | money_variants(ds["summary"].get("funding_gap"), unit)))

    lines = ["# 数字一致性核对报告", "", f"正文：`{os.path.basename(args.doc)}`　模型输出：`{os.path.basename(args.outputs)}`", "", "## 一、必提数字"]
    missing = 0
    for label, variants in required:
        hit = find_in_text(text, variants)
        if hit:
            lines.append(f"- ✅ {label}：正文出现 `{hit}`")
        else:
            missing += 1
            ex = sorted((t[0] + ("亿" if t[1] is True else (" million" if t[1] == "million" else "")) for t in variants if t[1] is not True), key=len)[:3]
            lines.append(f"- ❌ {label}：正文未出现（模型值形如 {', '.join(ex)}）")

    # 可疑数字扫描：限定在财务相关章节
    money_vals, pct_vals, ratio_vals = collect_model_numbers(o)
    for extra in args.extra:
        collect_extra_numbers(extra, money_vals, pct_vals, ratio_vals)
    chapters = re.split(r"\n(?=#\s)", text) if args.doc.endswith(".md") else [text]
    fin_text = []
    for ch in chapters:
        head = ch.strip().splitlines()[0] if ch.strip() else ""
        if any(k in head for k in ("财务", "融资", "资金", "执行摘要", "情景", "敏感", "贷款", "偿债", "还款", "申请", "筹措", "效益", "决策", "方案比较", "Financial", "Funding", "Executive", "Loan", "Debt", "Repayment", "Benefit", "Decision")) or len(chapters) == 1:
            fin_text.append(ch)
    scan = "\n".join(fin_text)
    suspicious = []
    checked = 0
    sourced = 0
    for m in NUM_TOKEN_RE.finditer(scan):
        raw, u = m.group(1), m.group(2)
        try:
            v = float(raw.replace(",", ""))
        except ValueError:
            continue
        checked += 1
        # 紧跟来源/登记册标记的数字视为已有出处
        tail = scan[m.end():m.end() + 45]
        if SOURCED_RE.search(tail) and not re.search(r"[。；\n]", tail.split("[")[0]):
            sourced += 1
            continue
        ok = False
        if u in ("%", "％"):
            ok = any(abs(v / 100 - p) <= max(args.tolerance * abs(p), 0.0051) for p in pct_vals if p is not None)
        elif u in ("倍", "x"):
            ok = any(abs(v - r) <= 0.011 for r in ratio_vals if r is not None)
        else:
            ul = u.lower()
            if ul in ("亿元", "亿") and unit == "万元":
                cand = v * 10000
            elif ul == "million" and unit == "万元":
                cand = v * 100
            elif ul == "billion" and unit == "万元":
                cand = v * 100000
            else:
                cand = v
            ok = any(abs(cand - mv) <= max(args.tolerance * abs(mv), 0.6) for mv in money_vals if mv is not None)
        if not ok:
            ctx_start = max(0, m.start() - 25)
            ctx = scan[ctx_start:m.end() + 15].replace("\n", " ")
            suspicious.append(f"`{raw}{u}` … {ctx}")
    lines += ["", "## 二、财务相关章节中与模型对不上的数字（需复核）", "",
              f"扫描了 {checked} 个带单位的数字，其中 {sourced} 个紧跟来源/登记册标记已跳过。下列数字在模型输出（含派生值与 --extra 文件）中找不到对应值（容差 {args.tolerance:.1%}）。来自行业数据的数字属正常，但预测类数字必须来自模型："]
    if suspicious:
        seen = set()
        for sline in suspicious:
            key = sline.split(" … ")[0]
            if key in seen:
                continue
            seen.add(key)
            lines.append(f"- {sline}")
    else:
        lines.append("- （无）")
    lines += ["", "## 三、结论", "", ("✅ 必提数字齐全。" if missing == 0 else f"❌ 有 {missing} 项必提数字缺失，请补充或核对口径。") + (f" 另有 {len(suspicious)} 处数字需人工复核。" if suspicious else "")]
    report = "\n".join(lines)
    with open(args.report, "w", encoding="utf-8") as f:
        f.write(report)
    print(report)
    sys.exit(1 if missing else 0)


if __name__ == "__main__":
    main()
