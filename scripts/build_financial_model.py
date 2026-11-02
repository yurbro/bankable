#!/usr/bin/env python3
"""
build_financial_model.py — 从 assumptions.json 生成带公式的 Excel 三表财务模型。

用法:
    python build_financial_model.py assumptions.json -o 模型.xlsx [--lang en]
    --lang en：工作表名与所有固定标签为英文（用户在假设文件里写的名称、来源保持原样），公式引用自动改写。

生成的工作簿（全部公式可追溯，蓝色字=输入，黑色=公式）:
    说明 | 假设 | 收入 | 费用 | 投资与折旧 | 融资 | 利润表 | 现金流量表 | 资产负债表 | 指标 | 情景对比 | 月度现金流(第N年) | 来源日志 | _map(隐藏)

_map 工作表记录每个关键行的位置，供 recalc_and_extract.py 读取。
假设文件的字段说明见 references/financial-model-guide.md，示例见 assets/assumptions-example.json。
"""
import argparse
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# 样式
# ----------------------------------------------------------------------------
MONEY = '#,##0.0;[Red]-#,##0.0;"-"'
MONEY2 = '#,##0.00;[Red]-#,##0.00;"-"'
INT = '#,##0;[Red]-#,##0;"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
RATIO = '0.00;[Red]-0.00;"-"'

F_INPUT = Font(name="Calibri", size=10, color="0000FF")
F_CALC = Font(name="Calibri", size=10, color="000000")
F_BOLD = Font(name="Calibri", size=10, bold=True)
F_TITLE = Font(name="Calibri", size=14, bold=True)
F_SECTION = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
F_NOTE = Font(name="Calibri", size=9, italic=True, color="666666")
FILL_SECTION = PatternFill("solid", fgColor="2F5496")
FILL_HEADER = PatternFill("solid", fgColor="D9E1F2")
FILL_TOTAL = PatternFill("solid", fgColor="F2F2F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF9E5")
THIN = Side(style="thin", color="BFBFBF")
BORDER_TOP = Border(top=Side(style="thin", color="000000"))
ALIGN_R = Alignment(horizontal="right")
ALIGN_C = Alignment(horizontal="center")

OPEX_CATEGORIES = ["销售费用", "管理费用", "研发费用", "其他费用"]
CATEGORY_ALIASES = {
    "销售": "销售费用", "销售费用": "销售费用", "市场": "销售费用", "营销": "销售费用",
    "管理": "管理费用", "管理费用": "管理费用", "行政": "管理费用",
    "研发": "研发费用", "研发费用": "研发费用", "技术": "研发费用",
    "其他": "其他费用", "其他费用": "其他费用",
    "营业成本": "营业成本", "生产": "营业成本", "成本": "营业成本",
    # English aliases (case-insensitive lookup below)
    "selling": "销售费用", "sales": "销售费用", "marketing": "销售费用", "sales & marketing": "销售费用", "s&m": "销售费用", "commercial": "销售费用",
    "g&a": "管理费用", "ga": "管理费用", "admin": "管理费用", "administration": "管理费用", "general": "管理费用", "management": "管理费用", "overheads": "管理费用", "overhead": "管理费用",
    "r&d": "研发费用", "rd": "研发费用", "research": "研发费用", "development": "研发费用", "technology": "研发费用",
    "other": "其他费用", "other opex": "其他费用", "operations": "其他费用",
    "cogs": "营业成本", "cost of sales": "营业成本", "production": "营业成本", "direct": "营业成本", "direct labour": "营业成本", "direct labor": "营业成本",
}


def category_of(value):
    key = str(value).strip()
    return CATEGORY_ALIASES.get(key) or CATEGORY_ALIASES.get(key.lower())


class ModelError(Exception):
    pass


# ----------------------------------------------------------------------------
# 假设文件校验与展开
# ----------------------------------------------------------------------------
def _err(errors, path, msg):
    errors.append(f"{path}: {msg}")


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def expand_growth(spec, n, path, errors):
    """把 {year1, growth} / {values} 规范化。

    返回 dict: {"mode": "growth", "year1": x, "growth": [g2..gn]} 或 {"mode": "values", "values": [...]}
    """
    if _num(spec):
        return {"mode": "values", "values": [float(spec)] * n}
    if isinstance(spec, list):
        if len(spec) != n:
            _err(errors, path, f"数组长度应为 {n}（预测年数），实际 {len(spec)}")
            return {"mode": "values", "values": [0.0] * n}
        if not all(_num(x) for x in spec):
            _err(errors, path, "数组元素必须是数字")
            return {"mode": "values", "values": [0.0] * n}
        return {"mode": "values", "values": [float(x) for x in spec]}
    if isinstance(spec, dict):
        if "values" in spec:
            return expand_growth(spec["values"], n, path + ".values", errors)
        if "year1" in spec:
            if not _num(spec["year1"]):
                _err(errors, path + ".year1", "必须是数字")
                return {"mode": "values", "values": [0.0] * n}
            g = spec.get("growth", 0.0)
            if _num(g):
                growth = [float(g)] * (n - 1)
            elif isinstance(g, list):
                if len(g) != n - 1:
                    _err(errors, path + ".growth", f"增长率数组长度应为 {n - 1}，实际 {len(g)}")
                    growth = [0.0] * (n - 1)
                else:
                    growth = [float(x) for x in g]
            else:
                _err(errors, path + ".growth", "必须是数字或数组")
                growth = [0.0] * (n - 1)
            return {"mode": "growth", "year1": float(spec["year1"]), "growth": growth}
    _err(errors, path, "格式应为数字、数组、{year1, growth} 或 {values}")
    return {"mode": "values", "values": [0.0] * n}


def validate(a):
    errors = []
    meta = a.get("meta")
    if not isinstance(meta, dict):
        raise ModelError("缺少 meta 块")
    n = meta.get("years")
    if not isinstance(n, int) or n < 3 or n > 10:
        _err(errors, "meta.years", "预测年数必须是 3–10 的整数")
        n = 5
    if not isinstance(meta.get("start_year"), int):
        _err(errors, "meta.start_year", "起始年份必须是整数，例如 2027")
    if not meta.get("project_name"):
        _err(errors, "meta.project_name", "请填写项目名称")

    rl = a.get("revenue_lines")
    if not isinstance(rl, list) or not rl:
        _err(errors, "revenue_lines", "至少需要一条收入线")
        rl = []
    for i, r in enumerate(rl):
        p = f"revenue_lines[{i}]"
        if not r.get("name"):
            _err(errors, p + ".name", "缺少名称")
        r["_volume"] = expand_growth(r.get("volume"), n, p + ".volume", errors)
        r["_price"] = expand_growth(r.get("price"), n, p + ".price", errors)
        if "unit_cost" in r:
            r["_cost"] = expand_growth(r["unit_cost"], n, p + ".unit_cost", errors)
            r["_cost_mode"] = "unit"
        elif "cogs_pct" in r:
            r["_cost"] = expand_growth(r["cogs_pct"], n, p + ".cogs_pct", errors)
            r["_cost_mode"] = "pct"
        else:
            _err(errors, p, "需要 unit_cost（单位成本）或 cogs_pct（成本率）之一")
            r["_cost"] = {"mode": "values", "values": [0.0] * n}
            r["_cost_mode"] = "pct"

    for i, o in enumerate(a.get("opex_lines", []) or []):
        p = f"opex_lines[{i}]"
        if not o.get("name"):
            _err(errors, p + ".name", "缺少名称")
        cat = category_of(o.get("category", "管理费用"))
        if cat is None or cat == "营业成本":
            _err(errors, p + ".category", f"费用类别应为 {OPEX_CATEGORIES} 之一（或英文 selling / g&a / r&d / other）")
            cat = "管理费用"
        o["_category"] = cat
        vals = o.get("values")
        if isinstance(vals, dict) and "pct_of_revenue" in vals:
            if not _num(vals["pct_of_revenue"]):
                _err(errors, p + ".values.pct_of_revenue", "必须是数字")
            o["_pct"] = float(vals.get("pct_of_revenue", 0))
        else:
            o["_series"] = expand_growth(vals, n, p + ".values", errors)

    for i, h in enumerate(a.get("headcount", []) or []):
        p = f"headcount[{i}]"
        if not h.get("role"):
            _err(errors, p + ".role", "缺少岗位名称")
        cnt = h.get("count")
        h["_count"] = expand_growth(cnt, n, p + ".count", errors)
        if not _num(h.get("annual_cost")):
            _err(errors, p + ".annual_cost", "人均年成本必须是数字（单位与 meta.unit 一致）")
        cat = category_of(h.get("category", "管理费用"))
        if cat is None:
            _err(errors, p + ".category", f"类别应为 {OPEX_CATEGORIES + ['营业成本']} 之一（或英文 selling / g&a / r&d / other / cogs）")
            cat = "管理费用"
        h["_category"] = cat

    for i, c in enumerate(a.get("capex", []) or []):
        p = f"capex[{i}]"
        if not c.get("name"):
            _err(errors, p + ".name", "缺少名称")
        if not _num(c.get("amount")):
            _err(errors, p + ".amount", "金额必须是数字")
        y = c.get("year")
        if not isinstance(y, int) or y < 1 or y > n:
            _err(errors, p + ".year", f"投入年份应为 1–{n} 的整数（第几个预测年）")
        if not _num(c.get("life_years")) or c.get("life_years", 0) <= 0:
            _err(errors, p + ".life_years", "折旧年限必须是正数")
        isy = c.get("in_service_year")
        if isy is not None and (not isinstance(isy, int) or (isinstance(y, int) and isy < y) or isy > n):
            _err(errors, p + ".in_service_year", f"投用年份应为 ≥ 投入年份且 ≤ {n} 的整数（建设期资产先计入在建工程，投用年起折旧）")

    wc = a.get("working_capital", {}) or {}
    for k in ("ar_days", "inventory_days", "ap_days"):
        if k in wc and not _num(wc[k]):
            _err(errors, f"working_capital.{k}", "必须是数字")

    fin = a.get("financing", {}) or {}
    for i, e in enumerate(fin.get("equity", []) or []):
        p = f"financing.equity[{i}]"
        if not isinstance(e.get("year"), int) or e["year"] < 1 or e["year"] > n:
            _err(errors, p + ".year", f"年份应为 1–{n} 的整数")
        if not _num(e.get("amount")):
            _err(errors, p + ".amount", "金额必须是数字")
    for i, g in enumerate(fin.get("grants", []) or []):
        p = f"financing.grants[{i}]"
        if not isinstance(g.get("year"), int) or g["year"] < 1 or g["year"] > n:
            _err(errors, p + ".year", f"年份应为 1–{n} 的整数")
        if not _num(g.get("amount")):
            _err(errors, p + ".amount", "金额必须是数字")
    for i, l in enumerate(fin.get("loans", []) or []):
        p = f"financing.loans[{i}]"
        for k in ("amount", "rate", "term_years"):
            if not _num(l.get(k)):
                _err(errors, f"{p}.{k}", "必须是数字")
        if not isinstance(l.get("year"), int) or l["year"] < 1 or l["year"] > n:
            _err(errors, p + ".year", f"提款年份应为 1–{n} 的整数")
        g = l.get("grace_years", 0)
        if not _num(g) or g < 0 or (_num(l.get("term_years")) and g >= l["term_years"]):
            _err(errors, p + ".grace_years", "宽限期必须 ≥0 且小于贷款期限")
        sch = l.get("schedule")
        if sch is not None:
            if not (isinstance(sch, list) and len(sch) == n and all(_num(x) for x in sch)):
                _err(errors, p + ".schedule", f"还本计划应为长度 {n} 的数组（各预测年偿还本金；到期一次还本则只在最后一年填金额）")
            elif _num(l.get("amount")) and sum(sch) - float(l["amount"]) > 0.5:
                _err(errors, p + ".schedule", f"还本计划合计 {sum(sch):,.1f} 超过贷款金额 {l['amount']:,.1f}（期限超出预测期时合计可小于贷款金额，余额留在期末）")
        dd = l.get("drawdown")
        if dd is not None:
            if not (isinstance(dd, list) and len(dd) == n and all(_num(x) for x in dd)):
                _err(errors, p + ".drawdown", f"提款计划应为长度 {n} 的数组（各预测年提款额）")
            elif _num(l.get("amount")) and abs(sum(dd) - float(l["amount"])) > 0.5:
                _err(errors, p + ".drawdown", f"提款计划合计 {sum(dd):,.1f} ≠ 贷款金额 {l['amount']:,.1f}")
            elif sch is None:
                _err(errors, p + ".drawdown", "分期提款时请同时给出 schedule（还本计划），等额本金规则不适用于分期提款")
    payout = fin.get("dividend_payout", 0)
    if not _num(payout) or payout < 0 or payout > 1:
        _err(errors, "financing.dividend_payout", "分红比例应在 0–1 之间")
    if fin.get("fixed_distributions") is not None:
        fin["_fixed_distributions"] = expand_growth(fin["fixed_distributions"], n, "financing.fixed_distributions", errors)

    tax = a.get("tax", {}) or {}
    if "income_tax_rate" in tax and not (_num(tax["income_tax_rate"]) and 0 <= tax["income_tax_rate"] < 1):
        _err(errors, "tax.income_tax_rate", "所得税率应在 0–1 之间，例如 0.25")

    for code in ("downside", "upside"):
        scn = (a.get("scenarios", {}) or {}).get(code, {}) or {}
        for key in ("volume", "price", "unit_cost", "opex"):
            v = scn.get(key)
            if v is None:
                continue
            if _num(v):
                scn[f"_{key}"] = [float(v)] * n
            elif isinstance(v, list) and len(v) == n and all(_num(x) for x in v):
                scn[f"_{key}"] = [float(x) for x in v]
            else:
                _err(errors, f"scenarios.{code}.{key}", f"应为数字或长度为 {n} 的数组")

    val = a.get("valuation", {}) or {}
    if "discount_rate" in val and not _num(val["discount_rate"]):
        _err(errors, "valuation.discount_rate", "折现率必须是数字")
    tm = val.get("terminal_method", "multiple")
    if tm not in ("multiple", "gordon", "none"):
        _err(errors, "valuation.terminal_method", "应为 multiple / gordon / none")

    ob = a.get("opening_balance")
    if meta.get("existing_business") and not isinstance(ob, dict):
        _err(errors, "opening_balance", "existing_business=true 时必须提供期初资产负债表")
    if isinstance(ob, dict):
        sch = ob.get("loan_schedule")
        if sch is not None:
            if not (isinstance(sch, list) and len(sch) == n and all(_num(x) for x in sch)):
                _err(errors, "opening_balance.loan_schedule", f"存量借款还本计划应为长度 {n} 的数组")
            elif sum(sch) - float(ob.get("loans", 0) or 0) > 0.5:
                _err(errors, "opening_balance.loan_schedule", f"还本计划合计 {sum(sch):,.1f} 超过存量借款余额 {float(ob.get('loans', 0) or 0):,.1f}")
        assets = sum(float(ob.get(k, 0) or 0) for k in ("cash", "accounts_receivable", "inventory", "other_current_assets", "fixed_assets_net", "land"))
        liab_eq = sum(float(ob.get(k, 0) or 0) for k in ("accounts_payable", "loans", "paid_in_capital", "capital_reserve", "retained_earnings"))
        if abs(assets - liab_eq) > 0.5:
            _err(errors, "opening_balance", f"期初资产负债表不平：资产 {assets:,.1f} ≠ 负债+权益 {liab_eq:,.1f}；可调整 retained_earnings 使其相等")

    m12 = a.get("monthly_year1")
    if isinstance(m12, dict):
        s = m12.get("seasonality")
        if s is not None:
            if not (isinstance(s, list) and len(s) == 12 and all(_num(x) for x in s)):
                _err(errors, "monthly_year1.seasonality", "应为 12 个数字")
            elif abs(sum(s) - 1.0) > 0.01:
                _err(errors, "monthly_year1.seasonality", f"12 个月占比之和应为 1，实际 {sum(s):.3f}")

    if errors:
        raise ModelError("假设文件有 %d 处问题:\n  - " % len(errors) + "\n  - ".join(errors))
    return n


# ----------------------------------------------------------------------------
# 工作表写入辅助
# ----------------------------------------------------------------------------
class Book:
    def __init__(self, n_years, start_year):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.n = n_years
        self.start_year = start_year
        self.cols = [get_column_letter(3 + i) for i in range(n_years)]
        self.map = []  # (key, sheet, row, col_start, n)
        self.sheets = {}

    def sheet(self, name, title, with_years=True):
        ws = self.wb.create_sheet(name)
        ws["A1"] = title
        ws["A1"].font = F_TITLE
        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 12
        if with_years:
            ws["A3"] = "年份"
            ws["A3"].font = F_BOLD
            ws["A4"] = "年序"
            ws["A4"].font = F_NOTE
            for i, c in enumerate(self.cols):
                ws[f"{c}3"] = self.start_year + i
                ws[f"{c}3"].font = F_BOLD
                ws[f"{c}3"].fill = FILL_HEADER
                ws[f"{c}3"].alignment = ALIGN_C
                ws[f"{c}4"] = i + 1
                ws[f"{c}4"].font = F_NOTE
                ws[f"{c}4"].alignment = ALIGN_C
                ws.column_dimensions[c].width = 13
            ws["B3"] = "单位"
            ws["B3"].font = F_BOLD
            ws["B3"].fill = FILL_HEADER
            ws.freeze_panes = "C5"
        self.sheets[name] = ws
        return Sheet(self, ws, name)

    def register(self, key, sheet, row, col_start=3, n=None):
        self.map.append((key, sheet, row, col_start, self.n if n is None else n))

    def ref(self, sheet, col, row, absolute=False):
        if absolute:
            return f"'{sheet}'!${col}${row}"
        return f"'{sheet}'!{col}{row}"

    def write_map(self):
        ws = self.wb.create_sheet("_map")
        ws.append(["key", "sheet", "row", "col_start", "n"])
        for rec in self.map:
            ws.append(list(rec))
        ws.sheet_state = "hidden"


class Sheet:
    def __init__(self, book, ws, name):
        self.book = book
        self.ws = ws
        self.name = name
        self.r = 6  # 当前行

    def section(self, title):
        ws = self.ws
        self.r += 1
        c = ws.cell(row=self.r, column=1, value=title)
        c.font = F_SECTION
        for col in range(1, 3 + self.book.n):
            ws.cell(row=self.r, column=col).fill = FILL_SECTION
        self.r += 1
        return self.r - 1

    def note(self, text):
        c = self.ws.cell(row=self.r, column=1, value=text)
        c.font = F_NOTE
        self.r += 1

    def blank(self):
        self.r += 1

    def row(self, label, cells, unit="", fmt=MONEY, key=None, kind="calc", bold=False, total=False):
        """写一行。cells: 长度 n 的列表，元素是数字(输入)或 '=...' 公式字符串或 None。"""
        ws = self.ws
        r = self.r
        ws.cell(row=r, column=1, value=label).font = F_BOLD if (bold or total) else F_CALC
        ws.cell(row=r, column=2, value=unit).font = F_NOTE
        for i, v in enumerate(cells):
            col = 3 + i
            c = ws.cell(row=r, column=col, value=v)
            c.number_format = fmt
            c.alignment = ALIGN_R
            if v is None:
                continue
            is_input = (kind == "input") and not (isinstance(v, str) and v.startswith("="))
            c.font = F_INPUT if is_input else (F_BOLD if (bold or total) else F_CALC)
            if is_input:
                c.fill = FILL_INPUT
            if total:
                c.fill = FILL_TOTAL
                c.border = BORDER_TOP
        if total:
            ws.cell(row=r, column=1).fill = FILL_TOTAL
        if key:
            self.book.register(key, self.name, r)
        self.r += 1
        return r

    def scalar(self, label, value, fmt=MONEY, key=None, kind="input", unit=""):
        """单个输入/公式单元格，写在 C 列。返回绝对引用。"""
        ws = self.ws
        r = self.r
        ws.cell(row=r, column=1, value=label).font = F_CALC
        ws.cell(row=r, column=2, value=unit).font = F_NOTE
        c = ws.cell(row=r, column=3, value=value)
        c.number_format = fmt
        c.alignment = ALIGN_R
        is_input = kind == "input" and not (isinstance(value, str) and str(value).startswith("="))
        c.font = F_INPUT if is_input else F_CALC
        if is_input:
            c.fill = FILL_INPUT
        if key:
            self.book.register(key, self.name, r, 3, 1)
        self.r += 1
        return f"'{self.name}'!$C${r}"

    def cell(self, col, row):
        return f"'{self.name}'!{col}{row}"

    def acell(self, col, row):
        return f"'{self.name}'!${col}${row}"


# ----------------------------------------------------------------------------
# 构建
# ----------------------------------------------------------------------------
def build(a, out_path, lang="zh"):
    n = validate(a)
    meta = a["meta"]
    unit = meta.get("unit", "万元")
    start_year = meta["start_year"]
    bk = Book(n, start_year)
    cols = bk.cols
    first, last = cols[0], cols[-1]
    yr = {}  # 单元格引用集合

    # ===================== 说明 =====================
    readme = bk.sheet("说明", f"{meta['project_name']} — 财务模型说明", with_years=False)
    ws = readme.ws
    lines = [
        f"生成日期：{date.today().isoformat()}    货币单位：{unit}（除注明外）    预测期：{start_year}–{start_year + n - 1}（{n} 年）",
        "",
        "结构：假设 → 收入 / 费用 / 投资与折旧 / 融资 → 利润表 → 现金流量表 → 资产负债表 → 指标 → 情景对比。",
        "颜色：蓝字黄底 = 输入（只改这些）；黑字 = 公式；灰底 = 合计。",
        "情景：在“假设”表的“情景选择”填 1（基准）、2（下行）、3（上行），全模型随之重算；“情景对比”表由脚本写入三情景的关键结果（静态值）。",
        "检查：资产负债表最后一行“检查差额”应为 0；现金流量表“期末现金”应始终 ≥ 最低现金要求，否则存在资金缺口。",
        "政府补助：按不占股的资本投入处理（计入资本公积与筹资现金流，不计税）；若为收益性补助或需按递延收益分摊，请在正文说明口径。",
        "简化处理：年度模型，期末计息、等额本金还款；所得税按利润总额×税率，亏损可结转抵扣（未设 5 年期限）；不含增值税与附加税；折旧直线法；营运资金按周转天数估算。",
        "免责：本模型基于假设表中的假设，用于商业计划分析，不构成审计意见、税务或法律建议，也不构成收益或融资承诺。",
        "",
        "由 Bankable skill 的 build_financial_model.py 生成。",
    ]
    for i, t in enumerate(lines):
        ws.cell(row=3 + i, column=1, value=t)
    ws.column_dimensions["A"].width = 120

    # ===================== 假设 =====================
    S = bk.sheet("假设", "假设表（蓝字为输入）")
    ws = S.ws
    S.r = 6
    S.section("基本信息")
    S.scalar("项目名称", meta["project_name"], fmt="@", key="project_name")
    S.scalar("货币单位", unit, fmt="@", key="unit")
    S.scalar("起始年份", start_year, fmt="0", key="start_year")
    S.scalar("预测年数", n, fmt="0", key="n_years")
    S.blank()

    S.section("情景")
    sel_ref = S.scalar("情景选择（1=基准 2=下行 3=上行）", 1, fmt="0", key="scenario_selector")
    sc = a.get("scenarios", {}) or {}
    dn = sc.get("downside", {}) or {}
    up = sc.get("upside", {}) or {}
    S.note("情景乘数按年填写（1 = 不变）；周期性行业可表达“低谷两年后回升”之类的路径。基准情景恒为 1。")
    mult = {}
    for key, label, dflt_dn, dflt_up in [
        ("volume", "销量乘数", 0.8, 1.15),
        ("price", "价格乘数", 0.95, 1.03),
        ("unit_cost", "单位成本乘数", 1.05, 0.98),
        ("opex", "费用乘数", 1.05, 1.0),
    ]:
        dn_vals = dn.get(f"_{key}") or [float(dn.get(key, dflt_dn))] * n
        up_vals = up.get(f"_{key}") or [float(up.get(key, dflt_up))] * n
        r_dn = S.row(f"{label} — 下行", dn_vals, unit="x", fmt=RATIO, kind="input", key=f"mult_{key}_downside")
        r_up = S.row(f"{label} — 上行", up_vals, unit="x", fmt=RATIO, kind="input", key=f"mult_{key}_upside")
        r_cur = S.row(f"{label} — 当前生效", [f"=CHOOSE({sel_ref},1,{cols[i]}{r_dn},{cols[i]}{r_up})" for i in range(n)], unit="x", fmt=RATIO, bold=True, key=f"mult_{key}")
        mult[key] = [f"'假设'!${cols[i]}${r_cur}" for i in range(n)]
    if sc.get("rationale"):
        S.note("情景依据：" + str(sc["rationale"]))
    if dn.get("label"):
        S.note("下行：" + str(dn["label"]))
    if up.get("label"):
        S.note("上行：" + str(up["label"]))
    S.blank()

    # ---- 收入假设
    S.section("收入假设（情景前的基准输入）")
    rev_refs = []
    for r_line in a["revenue_lines"]:
        name = r_line["name"]
        u = r_line.get("unit", "")
        S.note(f"{r_line.get('id', '')} {name}" + (f"（依据：{r_line['source']}）" if r_line.get("source") else ""))
        refs = {}
        for key, label, spec, fmt in [
            ("volume", "销量", r_line["_volume"], INT if u else MONEY),
            ("price", "单价", r_line["_price"], MONEY2),
        ]:
            refs[key] = write_series_input(S, label, spec, unit=(u if key == "volume" else f"{unit}/{u}" if u else unit), fmt=fmt)
        if r_line["_cost_mode"] == "unit":
            refs["cost"] = write_series_input(S, "单位成本", r_line["_cost"], unit=f"{unit}/{u}" if u else unit, fmt=MONEY2)
        else:
            refs["cost"] = write_series_input(S, "成本率（占收入）", r_line["_cost"], unit="%", fmt=PCT)
        refs["mode"] = r_line["_cost_mode"]
        refs["name"] = name
        refs["unit"] = u
        rev_refs.append(refs)
        S.blank()

    # ---- 费用假设
    S.section("费用假设（情景前的基准输入）")
    opex_refs = []
    for o in a.get("opex_lines", []) or []:
        label = f"{o['name']}（{o['_category']}）"
        if "_pct" in o:
            ref = write_series_input(S, label, {"mode": "values", "values": [o["_pct"]] * n}, unit="%收入", fmt=PCT)
            opex_refs.append({"name": o["name"], "category": o["_category"], "pct": True, "refs": ref})
        else:
            ref = write_series_input(S, label, o["_series"], unit=unit, fmt=MONEY)
            opex_refs.append({"name": o["name"], "category": o["_category"], "pct": False, "refs": ref})
        if o.get("source"):
            S.note("   依据：" + str(o["source"]))
    S.blank()

    S.section("人员假设")
    hc_refs = []
    for h in a.get("headcount", []) or []:
        cnt = write_series_input(S, f"{h['role']} — 人数（{h['_category']}）", h["_count"], unit="人", fmt=INT)
        cost = write_series_input(S, f"{h['role']} — 人均年成本", {"mode": "growth", "year1": float(h["annual_cost"]), "growth": [float(h.get("cost_growth", 0.0))] * (n - 1)}, unit=f"{unit}/人", fmt=MONEY2)
        hc_refs.append({"role": h["role"], "category": h["_category"], "count": cnt, "cost": cost})
    S.blank()

    S.section("资本开支假设")
    r = S.r
    for j, h in enumerate(["项目", "金额", "投入年序", "折旧年限", "残值率", "投用年序"]):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.font = F_BOLD
        c.fill = FILL_HEADER
    S.r += 1
    S.note("投用年序 = 开始折旧的年份；晚于投入年时，中间年份计入在建工程（建设期）。")
    capex_refs = []
    for c_item in a.get("capex", []) or []:
        r = S.r
        ws.cell(row=r, column=1, value=c_item["name"]).font = F_CALC
        isy = int(c_item.get("in_service_year") or c_item["year"])
        for j, (v, fmt) in enumerate([(float(c_item["amount"]), MONEY), (int(c_item["year"]), "0"), (float(c_item["life_years"]), "0.0"), (float(c_item.get("salvage_pct", 0.0)), PCT), (isy, "0")]):
            cell = ws.cell(row=r, column=2 + j, value=v)
            cell.number_format = fmt
            cell.font = F_INPUT
            cell.fill = FILL_INPUT
        capex_refs.append({"name": c_item["name"], "amount": f"'假设'!$B${r}", "year": f"'假设'!$C${r}", "life": f"'假设'!$D${r}", "salvage": f"'假设'!$E${r}", "in_service": f"'假设'!$F${r}"})
        S.r += 1
    S.blank()

    S.section("营运资金假设")
    wc = a.get("working_capital", {}) or {}
    ar_days = S.scalar("应收账款周转天数", float(wc.get("ar_days", 0)), fmt="0", unit="天", key="ar_days")
    inv_days = S.scalar("存货周转天数（按营业成本）", float(wc.get("inventory_days", 0)), fmt="0", unit="天", key="inventory_days")
    ap_days = S.scalar("应付账款周转天数（按成本+费用）", float(wc.get("ap_days", 0)), fmt="0", unit="天", key="ap_days")
    oca_pct = S.scalar("其他流动资产占收入比", float(wc.get("other_current_assets_pct", 0.0)), fmt=PCT, unit="%")
    S.blank()

    S.section("融资假设")
    fin = a.get("financing", {}) or {}
    r = S.r
    ws.cell(row=r, column=1, value="股权投入").font = F_BOLD
    for j, h in enumerate(["年序", "金额", "备注"]):
        c = ws.cell(row=r, column=2 + j, value=h)
        c.font = F_BOLD
        c.fill = FILL_HEADER
    S.r += 1
    eq_rows = []
    for e in fin.get("equity", []) or []:
        r = S.r
        ws.cell(row=r, column=1, value="股权投入").font = F_CALC
        c = ws.cell(row=r, column=2, value=int(e["year"])); c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = "0"
        c = ws.cell(row=r, column=3, value=float(e["amount"])); c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = MONEY
        ws.cell(row=r, column=4, value=e.get("note", "")).font = F_NOTE
        eq_rows.append(r)
        S.r += 1
    eq_range_year = f"'假设'!$B${eq_rows[0]}:$B${eq_rows[-1]}" if eq_rows else None
    eq_range_amt = f"'假设'!$C${eq_rows[0]}:$C${eq_rows[-1]}" if eq_rows else None
    S.blank()
    r = S.r
    ws.cell(row=r, column=1, value="政府补助 / 不占股资本投入").font = F_BOLD
    for j, h in enumerate(["年序", "金额", "备注"]):
        c = ws.cell(row=r, column=2 + j, value=h)
        c.font = F_BOLD
        c.fill = FILL_HEADER
    S.r += 1
    gr_rows = []
    for g in fin.get("grants", []) or []:
        r = S.r
        ws.cell(row=r, column=1, value="政府补助").font = F_CALC
        c = ws.cell(row=r, column=2, value=int(g["year"])); c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = "0"
        c = ws.cell(row=r, column=3, value=float(g["amount"])); c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = MONEY
        ws.cell(row=r, column=4, value=g.get("note", "")).font = F_NOTE
        gr_rows.append(r)
        S.r += 1
    gr_range_year = f"'假设'!$B${gr_rows[0]}:$B${gr_rows[-1]}" if gr_rows else None
    gr_range_amt = f"'假设'!$C${gr_rows[0]}:$C${gr_rows[-1]}" if gr_rows else None
    S.blank()
    r = S.r
    ws.cell(row=r, column=1, value="借款").font = F_BOLD
    for j, h in enumerate(["金额", "提款年序", "年利率", "期限(年)", "宽限期(年)"]):
        c = ws.cell(row=r, column=2 + j, value=h)
        c.font = F_BOLD
        c.fill = FILL_HEADER
    S.r += 1
    loan_refs = []
    for l in fin.get("loans", []) or []:
        r = S.r
        ws.cell(row=r, column=1, value=l.get("name", "借款")).font = F_CALC
        vals = [(float(l["amount"]), MONEY), (int(l["year"]), "0"), (float(l["rate"]), PCT), (float(l["term_years"]), "0"), (float(l.get("grace_years", 0)), "0")]
        for j, (v, fmt) in enumerate(vals):
            c = ws.cell(row=r, column=2 + j, value=v)
            c.font = F_INPUT; c.fill = FILL_INPUT; c.number_format = fmt
        loan_refs.append({"name": l.get("name", "借款"), "amount": f"'假设'!$B${r}", "year": f"'假设'!$C${r}", "rate": f"'假设'!$D${r}", "term": f"'假设'!$E${r}", "grace": f"'假设'!$F${r}", "schedule": None, "drawdown": None})
        S.r += 1
        if l.get("drawdown") is not None:
            rr = S.row(f"   {l.get('name', '借款')} — 提款计划（按年）", [float(x) for x in l["drawdown"]], unit=unit, fmt=MONEY, kind="input")
            loan_refs[-1]["drawdown"] = [f"'假设'!${cols[i]}${rr}" for i in range(n)]
        if l.get("schedule") is not None:
            rr = S.row(f"   {l.get('name', '借款')} — 还本计划（按年）", [float(x) for x in l["schedule"]], unit=unit, fmt=MONEY, kind="input")
            loan_refs[-1]["schedule"] = [f"'假设'!${cols[i]}${rr}" for i in range(n)]
    payout_ref = S.scalar("分红比例（占正净利润）", float(fin.get("dividend_payout", 0.0)), fmt=PCT, unit="%")
    if fin.get("_fixed_distributions"):
        fixed_dist_refs = write_series_input(S, "固定分配 / 合伙人提款（按年）", fin["_fixed_distributions"], unit=unit, fmt=MONEY)
    else:
        fixed_dist_refs = None
    S.blank()

    S.section("税务与估值假设")
    tax = a.get("tax", {}) or {}
    tax_ref = S.scalar("企业所得税率", float(tax.get("income_tax_rate", 0.25)), fmt=PCT, unit="%", key="tax_rate")
    loss0_ref = S.scalar("期初可抵扣亏损（负数）", -abs(float(tax.get("opening_loss_carryforward", 0.0))), fmt=MONEY, unit=unit)
    val = a.get("valuation", {}) or {}
    disc_ref = S.scalar("折现率（NPV）", float(val.get("discount_rate", 0.12)), fmt=PCT, unit="%", key="discount_rate")
    tm = val.get("terminal_method", "multiple")
    tm_code = {"none": 0, "multiple": 1, "gordon": 2}[tm]
    tm_ref = S.scalar("期末价值方法（0=无 1=EBITDA倍数 2=永续增长）", tm_code, fmt="0")
    ebm_ref = S.scalar("期末 EBITDA 倍数", float(val.get("ebitda_multiple", 8.0)), fmt=RATIO, unit="x")
    tg_ref = S.scalar("永续增长率", float(val.get("terminal_growth", 0.02)), fmt=PCT, unit="%")
    mincash_ref = S.scalar("最低现金要求", float(a.get("min_cash", 0.0)), fmt=MONEY, unit=unit, key="min_cash")
    S.blank()

    S.section("期初资产负债表（仅存量业务）")
    ob = a.get("opening_balance") if isinstance(a.get("opening_balance"), dict) else {}
    ob_ref = {}
    for key, label in [
        ("cash", "货币资金"), ("accounts_receivable", "应收账款"), ("inventory", "存货"), ("other_current_assets", "其他流动资产"),
        ("fixed_assets_net", "固定资产净值"), ("fixed_assets_remaining_life", "固定资产剩余折旧年限"), ("land", "土地 / 不折旧资产"),
        ("accounts_payable", "应付账款"), ("loans", "借款余额"), ("loan_rate", "存量借款年利率"), ("loan_remaining_years", "存量借款剩余年限"),
        ("paid_in_capital", "实收资本"), ("capital_reserve", "资本公积（含历史补助）"), ("retained_earnings", "未分配利润"),
    ]:
        fmt = PCT if key == "loan_rate" else ("0" if "life" in key or "years" in key else MONEY)
        ob_ref[key] = S.scalar(label, float(ob.get(key, 0) or 0), fmt=fmt, unit=("年" if ("life" in key or "years" in key) else ("%" if key == "loan_rate" else unit)), key=f"ob_{key}")
    ob_sched = None
    if isinstance(ob, dict) and ob.get("loan_schedule") is not None:
        rr = S.row("存量借款还本计划（按年）", [float(x) for x in ob["loan_schedule"]], unit=unit, fmt=MONEY, kind="input")
        ob_sched = [f"'假设'!${cols[i]}${rr}" for i in range(n)]
    ws.column_dimensions["A"].width = 40

    # ===================== 收入 =====================
    R = bk.sheet("收入", "收入与营业成本（已应用情景乘数）")
    rev_total_rows, cogs_total_rows, vol_rows = [], [], []
    for refs in rev_refs:
        R.section(refs["name"])
        u = refs["unit"]
        vrow = R.row("销量（情景后）", [f"={refs['volume'][i]}*{mult['volume'][i]}" for i in range(n)], unit=u, fmt=INT if u else MONEY, key=f"rev_volume_{refs['name']}")
        prow = R.row("单价（情景后）", [f"={refs['price'][i]}*{mult['price'][i]}" for i in range(n)], unit=f"{unit}/{u}" if u else unit, fmt=MONEY2)
        rrow = R.row("营业收入", [f"={cols[i]}{vrow}*{cols[i]}{prow}" for i in range(n)], unit=unit, bold=True, key=f"rev_{refs['name']}")
        if refs["mode"] == "unit":
            crow = R.row("单位成本（情景后）", [f"={refs['cost'][i]}*{mult['unit_cost'][i]}" for i in range(n)], unit=f"{unit}/{u}" if u else unit, fmt=MONEY2)
            corow = R.row("营业成本（材料/直接）", [f"={cols[i]}{vrow}*{cols[i]}{crow}" for i in range(n)], unit=unit, key=f"cogs_{refs['name']}")
        else:
            crow = R.row("成本率（情景后）", [f"={refs['cost'][i]}*{mult['unit_cost'][i]}" for i in range(n)], unit="%", fmt=PCT)
            corow = R.row("营业成本（材料/直接）", [f"={cols[i]}{rrow}*{cols[i]}{crow}" for i in range(n)], unit=unit, key=f"cogs_{refs['name']}")
        rev_total_rows.append(rrow)
        cogs_total_rows.append(corow)
        vol_rows.append(vrow)
        R.blank()
    R.section("合计")
    rev_total = R.row("营业收入合计", [f"={'+'.join(f'{cols[i]}{r}' for r in rev_total_rows)}" for i in range(n)], unit=unit, total=True, key="revenue")
    cogs_mat = R.row("营业成本（材料/直接）合计", [f"={'+'.join(f'{cols[i]}{r}' for r in cogs_total_rows)}" for i in range(n)], unit=unit, key="cogs_materials")
    # 生产人工 (from 费用 sheet, filled later) placeholder row
    cogs_labor_row = R.r
    R.row("营业成本（生产人工）", [None] * n, unit=unit, key="cogs_labor")
    cogs_total = R.row("营业成本合计", [f"={cols[i]}{cogs_mat}+{cols[i]}{cogs_labor_row}" for i in range(n)], unit=unit, total=True, key="cogs")
    gp = R.row("毛利", [f"={cols[i]}{rev_total}-{cols[i]}{cogs_total}" for i in range(n)], unit=unit, total=True, key="gross_profit")
    R.row("毛利率", [f"=IF({cols[i]}{rev_total}=0,0,{cols[i]}{gp}/{cols[i]}{rev_total})" for i in range(n)], unit="%", fmt=PCT, key="gross_margin")
    R.row("收入增长率", [None] + [f"=IF({cols[i-1]}{rev_total}=0,0,{cols[i]}{rev_total}/{cols[i-1]}{rev_total}-1)" for i in range(1, n)], unit="%", fmt=PCT, key="revenue_growth")
    yr["revenue"] = lambda i: f"'收入'!{cols[i]}{rev_total}"
    yr["cogs"] = lambda i: f"'收入'!{cols[i]}{cogs_total}"
    yr["gross_profit"] = lambda i: f"'收入'!{cols[i]}{gp}"

    # ===================== 费用 =====================
    E = bk.sheet("费用", "费用与人员成本（已应用情景乘数）")
    E.section("费用明细")
    cat_rows = {c: [] for c in OPEX_CATEGORIES}
    labor_rows = []
    for o in opex_refs:
        if o["pct"]:
            cells = [f"={yr['revenue'](i)}*{o['refs'][i]}*{mult['opex'][i]}" for i in range(n)]
        else:
            cells = [f"={o['refs'][i]}*{mult['opex'][i]}" for i in range(n)]
        rr = E.row(f"{o['name']}（{o['category']}）", cells, unit=unit)
        cat_rows[o["category"]].append(rr)
    E.blank()
    E.section("人员成本（人数 × 人均年成本）")
    for h in hc_refs:
        cells = [f"={h['count'][i]}*{h['cost'][i]}" + (f"*{mult['opex'][i]}" if h["category"] != "营业成本" else "") for i in range(n)]
        rr = E.row(f"{h['role']}（{h['category']}）", cells, unit=unit)
        if h["category"] == "营业成本":
            labor_rows.append(rr)
        else:
            cat_rows[h["category"]].append(rr)
    E.row("人员总数", [("=" + "+".join(h["count"][i] for h in hc_refs)) if hc_refs else 0 for i in range(n)], unit="人", fmt=INT, key="headcount_total")
    E.blank()
    E.section("分类合计")
    cat_total_rows = {}
    for cat in OPEX_CATEGORIES:
        rows = cat_rows[cat]
        cells = [("=" + "+".join(f"{cols[i]}{r}" for r in rows)) if rows else 0 for i in range(n)]
        cat_total_rows[cat] = E.row(cat, cells, unit=unit, total=True, key=f"opex_{cat}")
    opex_total = E.row("费用合计（不含折旧）", [f"={'+'.join(f'{cols[i]}{cat_total_rows[c]}' for c in OPEX_CATEGORIES)}" for i in range(n)], unit=unit, total=True, key="opex_total")
    labor_total = E.row("生产人工合计（计入营业成本）", [("=" + "+".join(f"{cols[i]}{r}" for r in labor_rows)) if labor_rows else 0 for i in range(n)], unit=unit, total=True, key="cogs_labor_total")
    # 回填收入表的生产人工行
    for i in range(n):
        R.ws[f"{cols[i]}{cogs_labor_row}"] = f"='费用'!{cols[i]}{labor_total}"
        R.ws[f"{cols[i]}{cogs_labor_row}"].number_format = MONEY
        R.ws[f"{cols[i]}{cogs_labor_row}"].font = F_CALC
    yr["opex"] = lambda i: f"'费用'!{cols[i]}{opex_total}"
    yr["opex_cat"] = {c: (lambda i, r=cat_total_rows[c]: f"'费用'!{cols[i]}{r}") for c in OPEX_CATEGORIES}

    # ===================== 投资与折旧 =====================
    C = bk.sheet("投资与折旧", "资本开支与折旧（直线法）")
    C.section("资本开支（投入）")
    capex_rows = []
    for c_item in capex_refs:
        rr = C.row(c_item["name"], [f"=IF({cols[i]}$4={c_item['year']},{c_item['amount']},0)" for i in range(n)], unit=unit)
        capex_rows.append(rr)
    capex_total = C.row("资本开支合计", [("=" + "+".join(f"{cols[i]}{r}" for r in capex_rows)) if capex_rows else 0 for i in range(n)], unit=unit, total=True, key="capex")
    C.blank()
    C.section("在建工程（投入后、投用前的资产）")
    cip_rows = []
    for c_item in capex_refs:
        f = [f"=IF(AND({cols[i]}$4>={c_item['year']},{cols[i]}$4<{c_item['in_service']}),{c_item['amount']},0)" for i in range(n)]
        cip_rows.append(C.row(f"在建：{c_item['name']}", f, unit=unit))
    cip_total = C.row("在建工程合计（年末）", [("=" + "+".join(f"{cols[i]}{r}" for r in cip_rows)) if cip_rows else 0 for i in range(n)], unit=unit, total=True, key="construction_in_progress")
    C.blank()
    C.section("折旧（自投用年起，直线法）")
    dep_rows = []
    for c_item in capex_refs:
        f = [f"=IF(AND({cols[i]}$4>={c_item['in_service']},{cols[i]}$4<{c_item['in_service']}+{c_item['life']}),{c_item['amount']}*(1-{c_item['salvage']})/{c_item['life']},0)" for i in range(n)]
        dep_rows.append(C.row(f"折旧：{c_item['name']}", f, unit=unit))
    dep_exist = C.row("折旧：期初固定资产", [f"=IF(AND({ob_ref['fixed_assets_remaining_life']}>0,{cols[i]}$4<={ob_ref['fixed_assets_remaining_life']}),{ob_ref['fixed_assets_net']}/{ob_ref['fixed_assets_remaining_life']},0)" for i in range(n)], unit=unit)
    dep_total = C.row("折旧与摊销合计", [f"={'+'.join(f'{cols[i]}{r}' for r in dep_rows + [dep_exist])}" for i in range(n)], unit=unit, total=True, key="depreciation")
    C.blank()
    C.section("固定资产")
    C.note("固定资产净值 = 期初 + 本年投入 − 本年折旧 − 在建工程年末余额变动（投用时从在建工程转入）。")
    fa_open = C.row("期初固定资产净值（含在建）", [f"={ob_ref['fixed_assets_net']}"] + [None] * (n - 1), unit=unit)
    fa_gross_close = C.row("期末固定资产净值（含在建）", [f"={cols[i]}{fa_open}+{cols[i]}{capex_total}-{cols[i]}{dep_total}" for i in range(n)], unit=unit, total=True, key="fixed_assets_incl_cip")
    fa_close = C.row("其中：已投用固定资产净值", [f"={cols[i]}{fa_gross_close}-{cols[i]}{cip_total}" for i in range(n)], unit=unit, key="fixed_assets_net")
    for i in range(1, n):
        C.ws[f"{cols[i]}{fa_open}"] = f"={cols[i-1]}{fa_gross_close}"
        C.ws[f"{cols[i]}{fa_open}"].number_format = MONEY
        C.ws[f"{cols[i]}{fa_open}"].font = F_CALC
    yr["capex"] = lambda i: f"'投资与折旧'!{cols[i]}{capex_total}"
    yr["dep"] = lambda i: f"'投资与折旧'!{cols[i]}{dep_total}"
    yr["fa"] = lambda i: f"'投资与折旧'!{cols[i]}{fa_close}"
    yr["cip"] = lambda i: f"'投资与折旧'!{cols[i]}{cip_total}"

    # ===================== 融资 =====================
    Fn = bk.sheet("融资", "股权与借款（等额本金，期末计息）")
    Fn.section("股权")
    if eq_range_year:
        eq_row = Fn.row("股权投入", [f"=SUMIF({eq_range_year},{cols[i]}$4,{eq_range_amt})" for i in range(n)], unit=unit, key="equity_in")
    else:
        eq_row = Fn.row("股权投入", [0] * n, unit=unit, key="equity_in")
    if gr_range_year:
        gr_row = Fn.row("政府补助 / 不占股资本投入", [f"=SUMIF({gr_range_year},{cols[i]}$4,{gr_range_amt})" for i in range(n)], unit=unit, key="grant_in")
    else:
        gr_row = Fn.row("政府补助 / 不占股资本投入", [0] * n, unit=unit, key="grant_in")
    Fn.blank()
    loan_blocks = []
    for l in loan_refs:
        Fn.section(f"借款：{l['name']}")
        if l.get("drawdown"):
            draw = Fn.row("提款（按计划）", [f"={l['drawdown'][i]}" for i in range(n)], unit=unit)
        else:
            draw = Fn.row("提款", [f"=IF({cols[i]}$4={l['year']},{l['amount']},0)" for i in range(n)], unit=unit)
        open_r = Fn.r
        Fn.row("期初余额", [None] * n, unit=unit)
        if l.get("schedule"):
            rep = Fn.row("偿还本金（按计划）", [f"=MIN({cols[i]}{open_r},{l['schedule'][i]})" for i in range(n)], unit=unit)
        else:
            rep = Fn.row("偿还本金", [f"=IF(AND({cols[i]}$4>={l['year']}+{l['grace']},{cols[i]}$4<={l['year']}+{l['term']}-1),MIN({cols[i]}{open_r},{l['amount']}/({l['term']}-{l['grace']})),0)" for i in range(n)], unit=unit)
        close_r = Fn.row("期末余额", [f"={cols[i]}{open_r}-{cols[i]}{rep}" for i in range(n)], unit=unit, total=True)
        intr = Fn.row("利息", [f"={cols[i]}{open_r}*{l['rate']}" for i in range(n)], unit=unit)
        for i in range(n):
            prev = "0" if i == 0 else f"{cols[i-1]}{close_r}"
            Fn.ws[f"{cols[i]}{open_r}"] = f"={prev}+{cols[i]}{draw}"
            Fn.ws[f"{cols[i]}{open_r}"].number_format = MONEY
            Fn.ws[f"{cols[i]}{open_r}"].font = F_CALC
        loan_blocks.append({"draw": draw, "rep": rep, "close": close_r, "int": intr})
        Fn.blank()
    # 存量借款
    Fn.section("存量借款（期初余额）")
    ex_open = Fn.r
    Fn.row("期初余额", [None] * n, unit=unit)
    if ob_sched:
        ex_rep = Fn.row("偿还本金（按计划）", [f"=MIN({cols[i]}{ex_open},{ob_sched[i]})" for i in range(n)], unit=unit)
    else:
        ex_rep = Fn.row("偿还本金", [f"=IF(AND({ob_ref['loan_remaining_years']}>0,{cols[i]}$4<={ob_ref['loan_remaining_years']}),MIN({cols[i]}{ex_open},{ob_ref['loans']}/{ob_ref['loan_remaining_years']}),0)" for i in range(n)], unit=unit)
    ex_close = Fn.row("期末余额", [f"={cols[i]}{ex_open}-{cols[i]}{ex_rep}" for i in range(n)], unit=unit, total=True)
    ex_int = Fn.row("利息", [f"={cols[i]}{ex_open}*{ob_ref['loan_rate']}" for i in range(n)], unit=unit)
    for i in range(n):
        prev = f"{ob_ref['loans']}" if i == 0 else f"{cols[i-1]}{ex_close}"
        Fn.ws[f"{cols[i]}{ex_open}"] = f"={prev}"
        Fn.ws[f"{cols[i]}{ex_open}"].number_format = MONEY
        Fn.ws[f"{cols[i]}{ex_open}"].font = F_CALC
    loan_blocks.append({"draw": None, "rep": ex_rep, "close": ex_close, "int": ex_int})
    Fn.blank()
    Fn.section("合计")
    def sum_blocks(field, i):
        parts = [f"{cols[i]}{b[field]}" for b in loan_blocks if b[field]]
        return "=" + "+".join(parts) if parts else 0

    draw_total = Fn.row("借款提款合计", [sum_blocks("draw", i) for i in range(n)], unit=unit, total=True, key="loan_drawdown")
    rep_total = Fn.row("偿还本金合计", [sum_blocks("rep", i) for i in range(n)], unit=unit, total=True, key="loan_repayment")
    int_total = Fn.row("利息合计", [sum_blocks("int", i) for i in range(n)], unit=unit, total=True, key="interest")
    debt_total = Fn.row("期末借款余额合计", [sum_blocks("close", i) for i in range(n)], unit=unit, total=True, key="debt_balance")
    ds_total = Fn.row("本息合计（偿债额）", [f"={cols[i]}{rep_total}+{cols[i]}{int_total}" for i in range(n)], unit=unit, total=True, key="debt_service")
    yr["equity_in"] = lambda i: f"'融资'!{cols[i]}{eq_row}"
    yr["grant_in"] = lambda i: f"'融资'!{cols[i]}{gr_row}"
    yr["draw"] = lambda i: f"'融资'!{cols[i]}{draw_total}"
    yr["rep"] = lambda i: f"'融资'!{cols[i]}{rep_total}"
    yr["int"] = lambda i: f"'融资'!{cols[i]}{int_total}"
    yr["debt"] = lambda i: f"'融资'!{cols[i]}{debt_total}"
    yr["ds"] = lambda i: f"'融资'!{cols[i]}{ds_total}"

    # ===================== 利润表 =====================
    P = bk.sheet("利润表", "利润表")
    P.r = 6
    p_rev = P.row("营业收入", [f"={yr['revenue'](i)}" for i in range(n)], unit=unit, bold=True, key="is_revenue")
    p_cogs = P.row("营业成本", [f"={yr['cogs'](i)}" for i in range(n)], unit=unit, key="is_cogs")
    p_gp = P.row("毛利", [f"={cols[i]}{p_rev}-{cols[i]}{p_cogs}" for i in range(n)], unit=unit, total=True, key="is_gross_profit")
    P.row("毛利率", [f"=IF({cols[i]}{p_rev}=0,0,{cols[i]}{p_gp}/{cols[i]}{p_rev})" for i in range(n)], unit="%", fmt=PCT, key="is_gross_margin")
    p_cat = {}
    for cat in OPEX_CATEGORIES:
        p_cat[cat] = P.row(cat, [f"={yr['opex_cat'][cat](i)}" for i in range(n)], unit=unit, key=f"is_{cat}")
    p_ebitda = P.row("EBITDA", [f"={cols[i]}{p_gp}-{'-'.join(f'{cols[i]}{p_cat[c]}' for c in OPEX_CATEGORIES)}" for i in range(n)], unit=unit, total=True, key="ebitda")
    P.row("EBITDA 率", [f"=IF({cols[i]}{p_rev}=0,0,{cols[i]}{p_ebitda}/{cols[i]}{p_rev})" for i in range(n)], unit="%", fmt=PCT, key="ebitda_margin")
    p_dep = P.row("折旧与摊销", [f"={yr['dep'](i)}" for i in range(n)], unit=unit, key="is_depreciation")
    p_ebit = P.row("EBIT（营业利润）", [f"={cols[i]}{p_ebitda}-{cols[i]}{p_dep}" for i in range(n)], unit=unit, total=True, key="ebit")
    p_int = P.row("利息费用", [f"={yr['int'](i)}" for i in range(n)], unit=unit, key="is_interest")
    p_ebt = P.row("利润总额", [f"={cols[i]}{p_ebit}-{cols[i]}{p_int}" for i in range(n)], unit=unit, total=True, key="ebt")
    p_pool = P.r
    P.row("可抵扣亏损池（辅助，负数）", [None] * n, unit=unit, key="loss_pool")
    p_taxable = P.row("应纳税所得额", [f"=MAX(0,{(loss0_ref if i == 0 else f'{cols[i-1]}{p_pool}')}+{cols[i]}{p_ebt})" for i in range(n)], unit=unit, key="taxable_income")
    for i in range(n):
        prev = loss0_ref if i == 0 else f"{cols[i-1]}{p_pool}"
        P.ws[f"{cols[i]}{p_pool}"] = f"=MIN(0,{prev}+{cols[i]}{p_ebt})"
        P.ws[f"{cols[i]}{p_pool}"].number_format = MONEY
        P.ws[f"{cols[i]}{p_pool}"].font = F_NOTE
    p_tax = P.row("所得税", [f"={cols[i]}{p_taxable}*{tax_ref}" for i in range(n)], unit=unit, key="tax")
    p_ni = P.row("净利润", [f"={cols[i]}{p_ebt}-{cols[i]}{p_tax}" for i in range(n)], unit=unit, total=True, key="net_income")
    P.row("净利率", [f"=IF({cols[i]}{p_rev}=0,0,{cols[i]}{p_ni}/{cols[i]}{p_rev})" for i in range(n)], unit="%", fmt=PCT, key="net_margin")
    p_div = P.row("分红 / 分配", [f"=MAX(0,{cols[i]}{p_ni})*{payout_ref}" + (f"+{fixed_dist_refs[i]}" if fixed_dist_refs else "") for i in range(n)], unit=unit, key="dividends")
    p_ret = P.row("留存收益增加", [f"={cols[i]}{p_ni}-{cols[i]}{p_div}" for i in range(n)], unit=unit, key="retained_addition")
    P.row("累计净利润", [f"=SUM(${first}{p_ni}:{cols[i]}{p_ni})" for i in range(n)], unit=unit, key="cumulative_net_income")
    yr["ni"] = lambda i: f"'利润表'!{cols[i]}{p_ni}"
    yr["ebitda"] = lambda i: f"'利润表'!{cols[i]}{p_ebitda}"
    yr["ebit"] = lambda i: f"'利润表'!{cols[i]}{p_ebit}"
    yr["tax"] = lambda i: f"'利润表'!{cols[i]}{p_tax}"
    yr["div"] = lambda i: f"'利润表'!{cols[i]}{p_div}"
    yr["ret"] = lambda i: f"'利润表'!{cols[i]}{p_ret}"
    yr["ebt"] = lambda i: f"'利润表'!{cols[i]}{p_ebt}"

    # ===================== 资产负债表（营运资金项先算） =====================
    B = bk.sheet("资产负债表", "资产负债表（年末）")
    B.r = 6
    B.section("资产")
    b_cash_r = B.r
    B.row("货币资金", [None] * n, unit=unit, key="cash")
    b_ar = B.row("应收账款", [f"={yr['revenue'](i)}*{ar_days}/365" for i in range(n)], unit=unit, key="accounts_receivable")
    b_inv = B.row("存货", [f"={yr['cogs'](i)}*{inv_days}/365" for i in range(n)], unit=unit, key="inventory")
    b_oca = B.row("其他流动资产", [f"={yr['revenue'](i)}*{oca_pct}" for i in range(n)], unit=unit, key="other_current_assets")
    b_ca = B.row("流动资产合计", [f"={cols[i]}{b_cash_r}+{cols[i]}{b_ar}+{cols[i]}{b_inv}+{cols[i]}{b_oca}" for i in range(n)], unit=unit, total=True, key="current_assets")
    b_fa = B.row("固定资产净值", [f"={yr['fa'](i)}" for i in range(n)], unit=unit, key="bs_fixed_assets")
    b_cip = B.row("在建工程", [f"={yr['cip'](i)}" for i in range(n)], unit=unit, key="bs_construction_in_progress")
    b_land = B.row("土地 / 不折旧资产", [f"={ob_ref['land']}" for i in range(n)], unit=unit, key="land")
    b_ta = B.row("资产总计", [f"={cols[i]}{b_ca}+{cols[i]}{b_fa}+{cols[i]}{b_cip}+{cols[i]}{b_land}" for i in range(n)], unit=unit, total=True, key="total_assets")
    B.blank()
    B.section("负债")
    b_ap = B.row("应付账款", [f"=({yr['cogs'](i)}+{yr['opex'](i)})*{ap_days}/365" for i in range(n)], unit=unit, key="accounts_payable")
    b_debt = B.row("借款余额", [f"={yr['debt'](i)}" for i in range(n)], unit=unit, key="bs_debt")
    b_tl = B.row("负债合计", [f"={cols[i]}{b_ap}+{cols[i]}{b_debt}" for i in range(n)], unit=unit, total=True, key="total_liabilities")
    B.blank()
    B.section("所有者权益")
    b_pic = B.row("实收资本", [f"={ob_ref['paid_in_capital']}+SUM('融资'!${first}{eq_row}:{cols[i]}{eq_row})" for i in range(n)], unit=unit, key="paid_in_capital")
    b_cap = B.row("资本公积（含政府补助）", [f"={ob_ref['capital_reserve']}+SUM('融资'!${first}{gr_row}:{cols[i]}{gr_row})" for i in range(n)], unit=unit, key="capital_reserve")
    b_re = B.row("未分配利润", [f"={ob_ref['retained_earnings']}+SUM('利润表'!${first}{p_ret}:{cols[i]}{p_ret})" for i in range(n)], unit=unit, key="retained_earnings")
    b_eq = B.row("所有者权益合计", [f"={cols[i]}{b_pic}+{cols[i]}{b_cap}+{cols[i]}{b_re}" for i in range(n)], unit=unit, total=True, key="total_equity")
    b_tle = B.row("负债和所有者权益总计", [f"={cols[i]}{b_tl}+{cols[i]}{b_eq}" for i in range(n)], unit=unit, total=True, key="total_liabilities_equity")
    B.blank()
    B.row("检查差额（应为 0）", [f"=ROUND({cols[i]}{b_ta}-{cols[i]}{b_tle},2)" for i in range(n)], unit=unit, fmt=MONEY2, key="balance_check", bold=True)
    B.row("资产负债率", [f"=IF({cols[i]}{b_ta}=0,0,{cols[i]}{b_tl}/{cols[i]}{b_ta})" for i in range(n)], unit="%", fmt=PCT, key="debt_ratio")
    b_cpltd = B.row("备查：一年内到期借款（次年应还本金）", [f"={yr['rep'](i + 1)}" if i < n - 1 else f"=MIN({yr['debt'](i)},{yr['rep'](i)})" for i in range(n)], unit=unit, key="current_portion_debt")
    B.row("流动比率（流动资产 ÷（应付账款 + 一年内到期借款））", [f"=IF({cols[i]}{b_ap}+{cols[i]}{b_cpltd}=0,\"-\",{cols[i]}{b_ca}/({cols[i]}{b_ap}+{cols[i]}{b_cpltd}))" for i in range(n)], unit="x", fmt=RATIO, key="current_ratio")

    # ===================== 现金流量表 =====================
    Cf = bk.sheet("现金流量表", "现金流量表（间接法）")
    Cf.r = 6
    Cf.section("经营活动")
    c_ni = Cf.row("净利润", [f"={yr['ni'](i)}" for i in range(n)], unit=unit, key="cf_net_income")
    c_dep = Cf.row("加：折旧与摊销", [f"={yr['dep'](i)}" for i in range(n)], unit=unit)
    def bs(i, row):
        return "'资产负债表'!" + cols[i] + str(row)

    def delta(sheet_row, opening_ref, sign=-1):
        """营运资金变动对现金的影响：资产增加为负，负债增加为正。"""
        out = []
        for i in range(n):
            prev = opening_ref if i == 0 else bs(i - 1, sheet_row)
            diff = f"({bs(i, sheet_row)}-{prev})"
            out.append(f"=-{diff}" if sign < 0 else f"={diff}")
        return out
    c_dar = Cf.row("减：应收账款增加", delta(b_ar, ob_ref["accounts_receivable"]), unit=unit)
    c_dinv = Cf.row("减：存货增加", delta(b_inv, ob_ref["inventory"]), unit=unit)
    c_doca = Cf.row("减：其他流动资产增加", delta(b_oca, ob_ref["other_current_assets"]), unit=unit)
    c_dap = Cf.row("加：应付账款增加", delta(b_ap, ob_ref["accounts_payable"], sign=+1), unit=unit)
    c_cfo = Cf.row("经营活动现金流净额", [f"=SUM({cols[i]}{c_ni}:{cols[i]}{c_dap})" for i in range(n)], unit=unit, total=True, key="cfo")
    Cf.blank()
    Cf.section("投资活动")
    c_capex = Cf.row("资本开支", [f"=-{yr['capex'](i)}" for i in range(n)], unit=unit)
    c_cfi = Cf.row("投资活动现金流净额", [f"={cols[i]}{c_capex}" for i in range(n)], unit=unit, total=True, key="cfi")
    Cf.blank()
    Cf.section("筹资活动")
    c_eq = Cf.row("股权投入", [f"={yr['equity_in'](i)}" for i in range(n)], unit=unit)
    c_gr = Cf.row("政府补助 / 不占股资本投入", [f"={yr['grant_in'](i)}" for i in range(n)], unit=unit)
    c_draw = Cf.row("借款提款", [f"={yr['draw'](i)}" for i in range(n)], unit=unit)
    c_rep = Cf.row("偿还本金", [f"=-{yr['rep'](i)}" for i in range(n)], unit=unit)
    c_div = Cf.row("分红", [f"=-{yr['div'](i)}" for i in range(n)], unit=unit)
    c_cff = Cf.row("筹资活动现金流净额", [f"=SUM({cols[i]}{c_eq}:{cols[i]}{c_div})" for i in range(n)], unit=unit, total=True, key="cff")
    Cf.blank()
    c_net = Cf.row("现金净增加额", [f"={cols[i]}{c_cfo}+{cols[i]}{c_cfi}+{cols[i]}{c_cff}" for i in range(n)], unit=unit, total=True, key="net_cash_change")
    c_open = Cf.row("期初现金", [f"={ob_ref['cash']}"] + [None] * (n - 1), unit=unit, key="opening_cash")
    c_close = Cf.row("期末现金", [f"={cols[i]}{c_open}+{cols[i]}{c_net}" for i in range(n)], unit=unit, total=True, key="closing_cash")
    for i in range(1, n):
        Cf.ws[f"{cols[i]}{c_open}"] = f"={cols[i-1]}{c_close}"
        Cf.ws[f"{cols[i]}{c_open}"].number_format = MONEY
        Cf.ws[f"{cols[i]}{c_open}"].font = F_CALC
    Cf.row("现金 − 最低现金要求（负数=缺口）", [f"={cols[i]}{c_close}-{mincash_ref}" for i in range(n)], unit=unit, key="cash_vs_min")
    Cf.blank()
    Cf.section("自由现金流（用于 IRR / NPV）")
    c_ufcf = Cf.row("无杠杆自由现金流", [f"={yr['ebit'](i)}*(1-{tax_ref})+{cols[i]}{c_dep}+{cols[i]}{c_dar}+{cols[i]}{c_dinv}+{cols[i]}{c_doca}+{cols[i]}{c_dap}+{cols[i]}{c_capex}" for i in range(n)], unit=unit, key="unlevered_fcf")
    Cf.row("累计无杠杆自由现金流", [f"=SUM(${first}{c_ufcf}:{cols[i]}{c_ufcf})" for i in range(n)], unit=unit, key="cumulative_ufcf")
    c_tv = Cf.row("期末价值（仅最后一年）", [None] * (n - 1) + [f"=CHOOSE({tm_ref}+1,0,{yr['ebitda'](n-1)}*{ebm_ref},IF({disc_ref}>{tg_ref},{last}{c_ufcf}*(1+{tg_ref})/({disc_ref}-{tg_ref}),0))"], unit=unit, key="terminal_value")
    c_ufcf_tv = Cf.row("无杠杆自由现金流 + 期末价值", [f"={cols[i]}{c_ufcf}+{cols[i]}{c_tv}" if i == n - 1 else f"={cols[i]}{c_ufcf}" for i in range(n)], unit=unit, key="ufcf_with_tv")
    c_eqcf = Cf.row("股东现金流（−投入 + 分红 + 期末权益价值）", [f"=-{cols[i]}{c_eq}-{cols[i]}{c_div}" + (f"+{cols[i]}{c_tv}-{yr['debt'](i)}+{cols[i]}{c_close}" if i == n - 1 else "") for i in range(n)], unit=unit, key="equity_cf")
    # 回填资产负债表现金
    for i in range(n):
        B.ws[f"{cols[i]}{b_cash_r}"] = f"='现金流量表'!{cols[i]}{c_close}"
        B.ws[f"{cols[i]}{b_cash_r}"].number_format = MONEY
        B.ws[f"{cols[i]}{b_cash_r}"].font = F_CALC

    # ===================== 指标 =====================
    K = bk.sheet("指标", "关键指标")
    K.r = 6
    K.section("逐年指标")
    K.row("收入增长率", [None] + [f"=IF({yr['revenue'](i-1)}=0,0,{yr['revenue'](i)}/{yr['revenue'](i-1)}-1)" for i in range(1, n)], unit="%", fmt=PCT, key="k_revenue_growth")
    K.row("毛利率", [f"=IF({yr['revenue'](i)}=0,0,{yr['gross_profit'](i)}/{yr['revenue'](i)})" for i in range(n)], unit="%", fmt=PCT, key="k_gross_margin")
    K.row("EBITDA 率", [f"=IF({yr['revenue'](i)}=0,0,{yr['ebitda'](i)}/{yr['revenue'](i)})" for i in range(n)], unit="%", fmt=PCT, key="k_ebitda_margin")
    K.row("净利率", [f"=IF({yr['revenue'](i)}=0,0,{yr['ni'](i)}/{yr['revenue'](i)})" for i in range(n)], unit="%", fmt=PCT, key="k_net_margin")
    K.row("净资产收益率（ROE）", [f"=IF('资产负债表'!{cols[i]}{b_eq}<=0,0,{yr['ni'](i)}/'资产负债表'!{cols[i]}{b_eq})" for i in range(n)], unit="%", fmt=PCT, key="roe")
    K.row("资产负债率", [f"='资产负债表'!{cols[i]}{b_tl}/'资产负债表'!{cols[i]}{b_ta}" for i in range(n)], unit="%", fmt=PCT, key="k_debt_ratio")
    K.row("利息保障倍数（EBIT/利息）", [f"=IF({yr['int'](i)}=0,\"-\",{yr['ebit'](i)}/{yr['int'](i)})" for i in range(n)], unit="x", fmt=RATIO, key="interest_coverage")
    k_dscr = K.row("偿债覆盖率 DSCR（(EBITDA−税)/本息）", [f"=IF({yr['ds'](i)}=0,\"-\",({yr['ebitda'](i)}-{yr['tax'](i)})/{yr['ds'](i)})" for i in range(n)], unit="x", fmt=RATIO, key="dscr")
    K.row("经营现金流 / 本息", [f"=IF({yr['ds'](i)}=0,\"-\",'现金流量表'!{cols[i]}{c_cfo}/{yr['ds'](i)})" for i in range(n)], unit="x", fmt=RATIO, key="cfo_to_debt_service")
    K.row("盈亏平衡收入（固定成本/毛利率）", [f"=IF({yr['gross_profit'](i)}<=0,\"-\",({yr['opex'](i)}+{yr['dep'](i)}+{yr['int'](i)})/({yr['gross_profit'](i)}/{yr['revenue'](i)}))" for i in range(n)], unit=unit, key="breakeven_revenue")
    K.row("盈亏平衡收入 / 实际收入", [f"=IF({yr['gross_profit'](i)}<=0,\"-\",({yr['opex'](i)}+{yr['dep'](i)}+{yr['int'](i)})/({yr['gross_profit'](i)}/{yr['revenue'](i)})/{yr['revenue'](i)})" for i in range(n)], unit="%", fmt=PCT, key="breakeven_ratio")
    K.row("期末现金", [f"='现金流量表'!{cols[i]}{c_close}" for i in range(n)], unit=unit, key="k_closing_cash")
    K.blank()
    K.section("汇总指标")
    ufcf_rng = f"'现金流量表'!{first}{c_ufcf}:{last}{c_ufcf}"
    ufcf_tv_rng = f"'现金流量表'!{first}{c_ufcf_tv}:{last}{c_ufcf_tv}"
    eqcf_rng = f"'现金流量表'!{first}{c_eqcf}:{last}{c_eqcf}"
    close_rng = f"'现金流量表'!{first}{c_close}:{last}{c_close}"
    K.scalar("项目 IRR（全投资口径，含期末价值）", f"=IFERROR(IRR({ufcf_tv_rng}),\"-\")", fmt=PCT, kind="calc", key="irr_with_tv")
    K.scalar("项目 IRR（全投资口径，不含期末价值）", f"=IFERROR(IRR({ufcf_rng}),\"-\")", fmt=PCT, kind="calc", key="irr_no_tv")
    K.scalar("项目 NPV（含期末价值）", f"=NPV({disc_ref},{ufcf_tv_rng})", fmt=MONEY, kind="calc", key="npv", unit=unit)
    K.scalar("股权 IRR（含期末权益价值）", f"=IFERROR(IRR({eqcf_rng}),\"-\")", fmt=PCT, kind="calc", key="equity_irr")
    K.scalar("最低期末现金", f"=MIN({close_rng})", fmt=MONEY, kind="calc", key="min_cash_balance", unit=unit)
    K.scalar("最低现金出现年份", f"=INDEX('现金流量表'!{first}$3:{last}$3,MATCH(MIN({close_rng}),{close_rng},0))", fmt="0", kind="calc", key="min_cash_year")
    K.scalar("资金缺口（最低现金要求 − 最低期末现金，≤0 表示无缺口）", f"=MAX(0,{mincash_ref}-MIN({close_rng}))", fmt=MONEY, kind="calc", key="funding_gap", unit=unit)
    K.scalar("股权投入合计", f"=SUM('融资'!{first}{eq_row}:{last}{eq_row})", fmt=MONEY, kind="calc", key="total_equity_in", unit=unit)
    K.scalar("借款提款合计", f"=SUM('融资'!{first}{draw_total}:{last}{draw_total})", fmt=MONEY, kind="calc", key="total_loan_drawdown", unit=unit)
    K.scalar("政府补助合计", f"=SUM('融资'!{first}{gr_row}:{last}{gr_row})", fmt=MONEY, kind="calc", key="total_grant_in", unit=unit)
    K.scalar("资本开支合计", f"=SUM('投资与折旧'!{first}{capex_total}:{last}{capex_total})", fmt=MONEY, kind="calc", key="total_capex", unit=unit)
    K.scalar("预测期累计净利润", f"=SUM('利润表'!{first}{p_ni}:{last}{p_ni})", fmt=MONEY, kind="calc", key="total_net_income", unit=unit)
    K.scalar("预测期累计经营现金流", f"=SUM('现金流量表'!{first}{c_cfo}:{last}{c_cfo})", fmt=MONEY, kind="calc", key="total_cfo", unit=unit)
    K.scalar("最低 DSCR（有偿债年份）", f"=IF(COUNT({first}{k_dscr}:{last}{k_dscr})=0,\"-\",MIN({first}{k_dscr}:{last}{k_dscr}))", fmt=RATIO, kind="calc", key="dscr_min")
    K.scalar("收入 CAGR（第1年→末年）", f"=IF({yr['revenue'](0)}<=0,\"-\",({yr['revenue'](n-1)}/{yr['revenue'](0)})^(1/{n-1})-1)", fmt=PCT, kind="calc", key="revenue_cagr")
    K.ws.column_dimensions["A"].width = 52

    # ===================== 情景对比（占位，由脚本写入） =====================
    Sc = bk.sheet("情景对比", "情景对比（由 recalc_and_extract.py 写入静态值）", with_years=False)
    Sc.ws["A3"] = "运行 scripts/recalc_and_extract.py 后此表将填入三种情景的关键结果。"
    Sc.ws["A3"].font = F_NOTE
    bk.register("scenario_table_anchor", "情景对比", 5, 1, 1)

    # ===================== 月度现金流（第1年） =====================
    m12 = a.get("monthly_year1")
    if isinstance(m12, dict):
        m_idx = int(m12.get("year_index", 1))
        m_idx = min(max(m_idx, 1), n)
        mi = m_idx - 1
        m_sheet = f"月度现金流(第{m_idx}年)"
        M = bk.sheet(m_sheet, f"第 {m_idx} 年（{start_year + mi}）月度现金流（参考口径）", with_years=False)
        ws = M.ws
        ws["A3"] = "月份"
        ws["A3"].font = F_BOLD
        mcols = [get_column_letter(3 + i) for i in range(12)]
        for i, c in enumerate(mcols):
            ws[f"{c}3"] = i + 1
            ws[f"{c}3"].font = F_BOLD
            ws[f"{c}3"].fill = FILL_HEADER
            ws[f"{c}3"].alignment = ALIGN_C
            ws.column_dimensions[c].width = 11
        ws["B3"] = "单位"
        ws.column_dimensions["O"].width = 12
        ws["O3"] = "合计"
        ws["O3"].font = F_BOLD
        M.r = 5
        seas = m12.get("seasonality") or [1 / 12] * 12
        # 收付款滞后默认由周转天数推出（与年度表口径一致）
        clag = int(m12["collection_lag_months"]) if m12.get("collection_lag_months") is not None else min(6, int(round(float(wc.get("ar_days", 30)) / 30)))
        plag = int(m12["payment_lag_months"]) if m12.get("payment_lag_months") is not None else min(6, int(round(float(wc.get("ap_days", 30)) / 30)))
        eq_m = int(m12.get("equity_month", 1))
        loan_m = int(m12.get("loan_month", 1))
        capex_months = m12.get("capex_months", {}) or {}
        M.note("说明：收入按季节性占比分摊并按收款滞后月数收款；材料成本按季节性分摊并按付款滞后付款；费用与人工按月平均；利息自提款月起平均计提；本金与所得税年末支付。")
        M.note("口径差异：年度表按周转天数计算年末应收/应付/存货，月度表按整月滞后且不含存货占用，因此 12 月末现金与年度表期末现金会有差异（见表末“与年度表差异”行）；月度表用于看现金节奏与最低点，年度表为正式口径。")
        r_seas = M.r
        ws.cell(row=r_seas, column=1, value="季节性占比").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_seas}"]
            cell.value = float(seas[i]); cell.number_format = PCT; cell.font = F_INPUT; cell.fill = FILL_INPUT
        M.r += 1
        lag_c = M.scalar("收款滞后（月）", clag, fmt="0")
        lag_p = M.scalar("付款滞后（月）", plag, fmt="0")
        M.blank()
        M.section("现金流入")
        y1 = lambda key: yr[key](mi)
        mcol = cols[mi]
        r_rec = M.r
        ws.cell(row=r_rec, column=1, value="销售收款").font = F_CALC
        for i, c in enumerate(mcols):
            # 收款 = sum over k: revenue × seas[k] where k + lag == i
            cell = ws[f"{c}{r_rec}"]
            cell.value = f"=IF({i+1}-{lag_c}>=1,{y1('revenue')}*INDEX(${mcols[0]}${r_seas}:${mcols[-1]}${r_seas},1,{i+1}-{lag_c}),0)"
            cell.number_format = MONEY; cell.font = F_CALC
        M.r += 1
        r_eq = M.r
        ws.cell(row=r_eq, column=1, value="股权投入").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_eq}"]; cell.value = f"=IF({c}$3={eq_m},{y1('equity_in')},0)"; cell.number_format = MONEY
        M.r += 1
        gr_m = int(m12.get("grant_month", 12))
        r_gr = M.r
        ws.cell(row=r_gr, column=1, value="政府补助").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_gr}"]; cell.value = f"=IF({c}$3={gr_m},{y1('grant_in')},0)"; cell.number_format = MONEY
        M.r += 1
        r_ln = M.r
        ws.cell(row=r_ln, column=1, value="借款提款").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_ln}"]; cell.value = f"=IF({c}$3={loan_m},{y1('draw')},0)"; cell.number_format = MONEY
        M.r += 1
        r_in = M.r
        ws.cell(row=r_in, column=1, value="流入合计").font = F_BOLD
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_in}"]; cell.value = f"=SUM({c}{r_rec}:{c}{r_ln})"; cell.number_format = MONEY; cell.font = F_BOLD; cell.fill = FILL_TOTAL
        M.r += 1
        M.blank()
        M.section("现金流出")
        r_mat = M.r
        ws.cell(row=r_mat, column=1, value="材料/直接成本付款").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_mat}"]
            cell.value = f"=IF({i+1}-{lag_p}>=1,'收入'!{mcol}{cogs_mat}*INDEX(${mcols[0]}${r_seas}:${mcols[-1]}${r_seas},1,{i+1}-{lag_p}),0)"
            cell.number_format = MONEY
        M.r += 1
        r_opx = M.r
        ws.cell(row=r_opx, column=1, value="费用与人工（含生产人工）").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_opx}"]; cell.value = f"=({y1('opex')}+'费用'!{mcol}{labor_total})/12"; cell.number_format = MONEY
        M.r += 1
        r_cx = M.r
        ws.cell(row=r_cx, column=1, value="资本开支").font = F_CALC
        for i, c in enumerate(mcols):
            parts = []
            for c_item, cref in zip(a.get("capex", []) or [], capex_refs):
                mth = int(capex_months.get(c_item.get("id", c_item["name"]), capex_months.get(c_item["name"], 1)))
                parts.append(f"IF(AND({cref['year']}={m_idx},{c}$3={mth}),{cref['amount']},0)")
            cell = ws[f"{c}{r_cx}"]; cell.value = "=" + ("+".join(parts) if parts else "0"); cell.number_format = MONEY
        M.r += 1
        r_int = M.r
        ws.cell(row=r_int, column=1, value="利息").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_int}"]; cell.value = f"=IF({c}$3>={loan_m},{y1('int')}/(13-{loan_m}),0)"; cell.number_format = MONEY
        M.r += 1
        r_rp = M.r
        ws.cell(row=r_rp, column=1, value="偿还本金（年末）").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_rp}"]; cell.value = f"=IF({c}$3=12,{y1('rep')},0)"; cell.number_format = MONEY
        M.r += 1
        r_tx = M.r
        ws.cell(row=r_tx, column=1, value="所得税（年末）").font = F_CALC
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_tx}"]; cell.value = f"=IF({c}$3=12,{y1('tax')},0)"; cell.number_format = MONEY
        M.r += 1
        r_out = M.r
        ws.cell(row=r_out, column=1, value="流出合计").font = F_BOLD
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_out}"]; cell.value = f"=SUM({c}{r_mat}:{c}{r_tx})"; cell.number_format = MONEY; cell.font = F_BOLD; cell.fill = FILL_TOTAL
        M.r += 1
        M.blank()
        r_netm = M.r
        ws.cell(row=r_netm, column=1, value="净现金流").font = F_BOLD
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_netm}"]; cell.value = f"={c}{r_in}-{c}{r_out}"; cell.number_format = MONEY; cell.font = F_BOLD
        M.r += 1
        r_openm = M.r
        ws.cell(row=r_openm, column=1, value="期初现金").font = F_CALC
        r_closem = r_openm + 1
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_openm}"]
            cell.value = (f"={ob_ref['cash']}" if mi == 0 else f"='现金流量表'!{cols[mi-1]}{c_close}") if i == 0 else f"={mcols[i-1]}{r_closem}"
            cell.number_format = MONEY
        M.r += 1
        ws.cell(row=r_closem, column=1, value="期末现金").font = F_BOLD
        for i, c in enumerate(mcols):
            cell = ws[f"{c}{r_closem}"]; cell.value = f"={c}{r_openm}+{c}{r_netm}"; cell.number_format = MONEY; cell.font = F_BOLD; cell.fill = FILL_TOTAL
        M.r += 1
        for rr in (r_rec, r_eq, r_gr, r_ln, r_in, r_mat, r_opx, r_cx, r_int, r_rp, r_tx, r_out, r_netm):
            cell = ws[f"O{rr}"]; cell.value = f"=SUM({mcols[0]}{rr}:{mcols[-1]}{rr})"; cell.number_format = MONEY; cell.font = F_BOLD
        M.blank()
        r_ann = M.r
        ws.cell(row=r_ann, column=1, value=f"年度表期末现金（第 {m_idx} 年）").font = F_CALC
        cell = ws[f"{mcols[-1]}{r_ann}"]; cell.value = f"='现金流量表'!{mcol}{c_close}"; cell.number_format = MONEY
        M.r += 1
        r_diff = M.r
        ws.cell(row=r_diff, column=1, value="与年度表差异（月度 − 年度，口径差异）").font = F_NOTE
        cell = ws[f"{mcols[-1]}{r_diff}"]; cell.value = f"={mcols[-1]}{r_closem}-{mcols[-1]}{r_ann}"; cell.number_format = MONEY; cell.font = F_NOTE
        M.r += 1
        bk.register("monthly_closing_cash", m_sheet, r_closem, 3, 12)
        bk.register("monthly_net_cash", m_sheet, r_netm, 3, 12)
        bk.register("monthly_collection_lag", m_sheet, int(lag_c.split("$")[-1]), 3, 1)
        bk.register("monthly_payment_lag", m_sheet, int(lag_p.split("$")[-1]), 3, 1)
        bk.register("monthly_year_index", m_sheet, 1, 1, 1)
        ws["A1"].value = f"第 {m_idx} 年（{start_year + mi}）月度现金流（参考口径）"
        ws.cell(row=2, column=1, value=m_idx).font = F_NOTE
        bk.map[-1] = ("monthly_year_index", m_sheet, 2, 1, 1)
        ws.column_dimensions["A"].width = 30

    # ===================== 来源日志 =====================
    L = bk.sheet("来源日志", "假设来源与证据日志", with_years=False)
    ws = L.ws
    headers = ["编号", "内容", "来源 / 依据", "等级", "置信度", "影响", "需确认"]
    for j, h in enumerate(headers):
        c = ws.cell(row=3, column=1 + j, value=h)
        c.font = F_BOLD; c.fill = FILL_HEADER
    r = 4
    for s in a.get("sources", []) or []:
        for j, k in enumerate(["id", "content", "source", "grade", "confidence", "impact", "confirm"]):
            ws.cell(row=r, column=1 + j, value=str(s.get(k, "")))
        r += 1
    # 自动汇总各块的 source 字段
    auto = []
    for rl in a["revenue_lines"]:
        if rl.get("source"):
            auto.append((rl.get("id", ""), f"收入线：{rl['name']}（销量/单价/成本）", rl["source"]))
    for o in a.get("opex_lines", []) or []:
        if o.get("source"):
            auto.append((o.get("id", ""), f"费用：{o['name']}", o["source"]))
    for c_item in a.get("capex", []) or []:
        if c_item.get("source"):
            auto.append((c_item.get("id", ""), f"资本开支：{c_item['name']}", c_item["source"]))
    for l in fin.get("loans", []) or []:
        if l.get("source"):
            auto.append((l.get("id", ""), f"借款：{l.get('name', '')}", l["source"]))
    if auto:
        r += 1
        ws.cell(row=r, column=1, value="模型字段引用的登记册编号").font = F_BOLD
        r += 1
        for rec in auto:
            ws.cell(row=r, column=1, value=rec[0]); ws.cell(row=r, column=2, value=rec[1]); ws.cell(row=r, column=3, value=rec[2])
            r += 1
    for col, w in zip("ABCDEFG", [10, 50, 50, 8, 10, 30, 10]):
        ws.column_dimensions[col].width = w

    bk.write_map()
    if lang == "en":
        from model_i18n import localize_workbook
        leftover = localize_workbook(bk.wb, unit=unit)
        if leftover:
            print(f"[提示] 英文模型中仍有 {leftover} 个含中文的单元格（通常是假设文件里用户自己写的名称或来源说明）")
    bk.wb.calculation.fullCalcOnLoad = True
    bk.wb.save(out_path)
    return out_path


def write_series_input(S, label, spec, unit="", fmt=MONEY):
    """在假设表写一个时间序列输入；growth 模式写 year1 + 增长率行（后续年公式）。返回每年的绝对引用列表。"""
    n = S.book.n
    cols = S.book.cols
    ws = S.ws
    if spec["mode"] == "values":
        r = S.row(label, spec["values"], unit=unit, fmt=fmt, kind="input")
        return [f"'假设'!${cols[i]}${r}" for i in range(n)]
    # growth 模式
    r = S.r
    g_r = r + 1
    cells = [spec["year1"]] + [f"={cols[i-1]}{r}*(1+{cols[i]}{g_r})" for i in range(1, n)]
    S.row(label, cells, unit=unit, fmt=fmt, kind="input")
    S.row("   增长率", [None] + list(spec["growth"]), unit="%", fmt=PCT, kind="input")
    ws.cell(row=g_r, column=1).font = F_NOTE
    return [f"'假设'!${cols[i]}${r}" for i in range(n)]


def main():
    ap = argparse.ArgumentParser(description="从 assumptions.json 生成 Excel 三表财务模型")
    ap.add_argument("assumptions")
    ap.add_argument("-o", "--output", default="财务模型.xlsx")
    ap.add_argument("--lang", default="zh", choices=["zh", "en"], help="工作表与标签语言（en = 英文模型，用户内容保持原样）")
    args = ap.parse_args()
    with open(args.assumptions, encoding="utf-8") as f:
        a = json.load(f)
    try:
        out = build(a, args.output, lang=args.lang)
    except ModelError as e:
        print("[错误] " + str(e), file=sys.stderr)
        sys.exit(1)
    print(f"[OK] 模型已生成：{out}")
    print("下一步：python scripts/recalc_and_extract.py " + out + " -o model_outputs.json")


if __name__ == "__main__":
    main()
