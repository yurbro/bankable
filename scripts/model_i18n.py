"""
model_i18n.py — 把 build_financial_model.py 生成的工作簿标签翻译成英文（--lang en）。

做法：先建好中文工作簿，再做一次本地化：重命名工作表并改写公式中的表引用，逐格翻译固定标签
（精确词典 + 模式规则），用户自己写的内容（收入线名称、来源说明等）保持原样。
recalc_and_extract.py 用 _map 定位行，不依赖工作表名，所以英文模型同样可重算。
"""
import re

SHEET_NAMES = {
    "说明": "Readme", "假设": "Assumptions", "收入": "Revenue", "费用": "Opex", "投资与折旧": "Capex & Depreciation",
    "融资": "Financing", "利润表": "Income Statement", "现金流量表": "Cash Flow", "资产负债表": "Balance Sheet",
    "指标": "Indicators", "情景对比": "Scenarios", "来源日志": "Source Log",
}
CATEGORIES = {"销售费用": "Selling", "管理费用": "G&A", "研发费用": "R&D", "其他费用": "Other opex", "营业成本": "COGS"}

EXACT = {
    # 假设表
    "假设表（蓝字为输入）": "Assumptions (blue = inputs)", "基本信息": "General", "项目名称": "Project", "货币单位": "Unit", "起始年份": "Start year", "预测年数": "Forecast years",
    "情景": "Scenarios", "情景选择（1=基准 2=下行 3=上行）": "Scenario selector (1 = base, 2 = downside, 3 = upside)",
    "情景乘数按年填写（1 = 不变）；周期性行业可表达“低谷两年后回升”之类的路径。基准情景恒为 1。": "Scenario multipliers by year (1 = unchanged); cyclical industries can model e.g. two trough years then recovery. Base case is always 1.",
    "销量乘数 — 下行": "Volume multiplier — downside", "销量乘数 — 上行": "Volume multiplier — upside", "销量乘数 — 当前生效": "Volume multiplier — active",
    "价格乘数 — 下行": "Price multiplier — downside", "价格乘数 — 上行": "Price multiplier — upside", "价格乘数 — 当前生效": "Price multiplier — active",
    "单位成本乘数 — 下行": "Unit-cost multiplier — downside", "单位成本乘数 — 上行": "Unit-cost multiplier — upside", "单位成本乘数 — 当前生效": "Unit-cost multiplier — active",
    "费用乘数 — 下行": "Opex multiplier — downside", "费用乘数 — 上行": "Opex multiplier — upside", "费用乘数 — 当前生效": "Opex multiplier — active",
    "收入假设（情景前的基准输入）": "Revenue assumptions (base inputs before scenario)", "销量": "Volume", "单价": "Price", "单位成本": "Unit cost", "成本率（占收入）": "Cost ratio (% of revenue)", "   增长率": "   growth",
    "费用假设（情景前的基准输入）": "Opex assumptions (base inputs before scenario)", "人员假设": "Headcount assumptions",
    "资本开支假设": "Capex assumptions", "项目": "Item", "金额": "Amount", "投入年序": "Spend year", "折旧年限": "Life (years)", "残值率": "Salvage %", "投用年序": "In-service year",
    "投用年序 = 开始折旧的年份；晚于投入年时，中间年份计入在建工程（建设期）。": "In-service year = first year of depreciation; before it the asset sits in construction in progress.",
    "营运资金假设": "Working-capital assumptions", "应收账款周转天数": "Receivable days", "存货周转天数（按营业成本）": "Inventory days (on COGS)", "应付账款周转天数（按成本+费用）": "Payable days (on COGS + opex)", "其他流动资产占收入比": "Other current assets (% of revenue)",
    "融资假设": "Financing assumptions", "股权投入": "Equity injected", "年序": "Year #", "备注": "Note", "政府补助 / 不占股资本投入": "Grants / non-equity capital", "政府补助": "Grant",
    "借款": "Loan", "提款年序": "Drawdown year", "年利率": "Interest rate", "期限(年)": "Term (yrs)", "宽限期(年)": "Grace (yrs)", "分红比例（占正净利润）": "Dividend payout (% of positive net income)",
    "税务与估值假设": "Tax and valuation assumptions", "企业所得税率": "Corporate income tax rate", "期初可抵扣亏损（负数）": "Opening loss carry-forward (negative)", "折现率（NPV）": "Discount rate (NPV)",
    "期末价值方法（0=无 1=EBITDA倍数 2=永续增长）": "Terminal value method (0 = none, 1 = EBITDA multiple, 2 = Gordon growth)", "期末 EBITDA 倍数": "Terminal EBITDA multiple", "永续增长率": "Terminal growth rate", "最低现金要求": "Minimum cash requirement",
    "期初资产负债表（仅存量业务）": "Opening balance sheet (existing business only)", "货币资金": "Cash", "应收账款": "Accounts receivable", "存货": "Inventory", "其他流动资产": "Other current assets",
    "固定资产净值": "Fixed assets, net", "固定资产剩余折旧年限": "Remaining life of fixed assets (yrs)", "应付账款": "Accounts payable", "借款余额": "Borrowings", "存量借款年利率": "Rate on existing loans", "存量借款剩余年限": "Remaining term of existing loans (yrs)",
    "实收资本": "Paid-in capital", "土地 / 不折旧资产": "Land / non-depreciating assets", "固定分配 / 合伙人提款（按年）": "Fixed distributions / partners' drawings (by year)", "分红 / 分配": "Dividends / distributions", "资本公积（含历史补助）": "Capital reserve (incl. past grants)", "未分配利润": "Retained earnings", "存量借款还本计划（按年）": "Existing-loan repayment schedule (by year)",
    "天": "days", "%收入": "% rev", "人": "FTE", "年": "yrs", "x": "x", "%": "%", "年份": "Year", "年序": "Year #",
    # 收入
    "收入与营业成本（已应用情景乘数）": "Revenue and cost of goods sold (scenario applied)", "销量（情景后）": "Volume (after scenario)", "单价（情景后）": "Price (after scenario)", "营业收入": "Revenue",
    "单位成本（情景后）": "Unit cost (after scenario)", "成本率（情景后）": "Cost ratio (after scenario)", "营业成本（材料/直接）": "COGS (materials / direct)", "合计": "Total",
    "营业收入合计": "Total revenue", "营业成本（材料/直接）合计": "Total COGS (materials / direct)", "营业成本（生产人工）": "COGS (production labour)", "营业成本合计": "Total COGS", "毛利": "Gross profit", "毛利率": "Gross margin", "收入增长率": "Revenue growth",
    # 费用
    "费用与人员成本（已应用情景乘数）": "Operating expenses and headcount cost (scenario applied)", "费用明细": "Opex lines", "人员成本（人数 × 人均年成本）": "Headcount cost (FTE × annual cost per FTE)", "人员总数": "Total headcount",
    "分类合计": "Totals by category", "销售费用": "Selling expenses", "管理费用": "G&A expenses", "研发费用": "R&D expenses", "其他费用": "Other expenses", "费用合计（不含折旧）": "Total opex (excl. depreciation)", "生产人工合计（计入营业成本）": "Total production labour (in COGS)",
    # 投资与折旧
    "资本开支与折旧（直线法）": "Capital expenditure and depreciation (straight line)", "资本开支（投入）": "Capital expenditure (spend)", "资本开支合计": "Total capex", "在建工程（投入后、投用前的资产）": "Construction in progress (spent, not yet in service)",
    "在建工程合计（年末）": "Total construction in progress (year-end)", "折旧（自投用年起，直线法）": "Depreciation (from in-service year, straight line)", "折旧：期初固定资产": "Depreciation: opening fixed assets", "折旧与摊销合计": "Total depreciation & amortisation", "固定资产": "Fixed assets",
    "固定资产净值 = 期初 + 本年投入 − 本年折旧 − 在建工程年末余额变动（投用时从在建工程转入）。": "Fixed assets, net = opening + spend − depreciation − change in construction in progress (transferred when placed in service).",
    "期初固定资产净值（含在建）": "Opening fixed assets incl. CIP", "期末固定资产净值（含在建）": "Closing fixed assets incl. CIP", "其中：已投用固定资产净值": "of which: fixed assets in service, net",
    # 融资
    "股权与借款（等额本金，期末计息）": "Equity and loans (equal principal, year-end interest)", "股权": "Equity", "提款": "Drawdown", "提款（按计划）": "Drawdown (scheduled)", "期初余额": "Opening balance", "偿还本金": "Principal repaid", "偿还本金（按计划）": "Principal repaid (scheduled)",
    "期末余额": "Closing balance", "利息": "Interest", "存量借款（期初余额）": "Existing loans (opening balance)", "借款提款合计": "Total loan drawdowns", "偿还本金合计": "Total principal repaid", "利息合计": "Total interest", "期末借款余额合计": "Total borrowings (closing)", "本息合计（偿债额）": "Debt service (principal + interest)",
    # 利润表
    "利润表": "Income statement", "营业成本": "Cost of goods sold", "EBITDA 率": "EBITDA margin", "折旧与摊销": "Depreciation & amortisation", "EBIT（营业利润）": "EBIT (operating profit)", "利息费用": "Interest expense", "利润总额": "Profit before tax",
    "可抵扣亏损池（辅助，负数）": "Loss carry-forward pool (helper, negative)", "应纳税所得额": "Taxable income", "所得税": "Income tax", "净利润": "Net income", "净利率": "Net margin", "分红": "Dividends", "留存收益增加": "Addition to retained earnings", "累计净利润": "Cumulative net income",
    # 资产负债表
    "资产负债表（年末）": "Balance sheet (year-end)", "资产": "Assets", "流动资产合计": "Total current assets", "在建工程": "Construction in progress", "资产总计": "Total assets", "负债": "Liabilities", "负债合计": "Total liabilities",
    "所有者权益": "Equity", "资本公积（含政府补助）": "Capital reserve (incl. grants)", "所有者权益合计": "Total equity", "负债和所有者权益总计": "Total liabilities and equity", "检查差额（应为 0）": "Balance check (must be 0)", "资产负债率": "Debt-to-asset ratio",
    "备查：一年内到期借款（次年应还本金）": "Memo: current portion of debt (next year's principal)", "流动比率（流动资产 ÷（应付账款 + 一年内到期借款））": "Current ratio (current assets ÷ (payables + current portion of debt))",
    # 现金流量表
    "现金流量表（间接法）": "Cash flow statement (indirect method)", "经营活动": "Operating activities", "加：折旧与摊销": "Add: depreciation & amortisation", "减：应收账款增加": "Less: increase in receivables", "减：存货增加": "Less: increase in inventory",
    "减：其他流动资产增加": "Less: increase in other current assets", "加：应付账款增加": "Add: increase in payables", "经营活动现金流净额": "Net cash from operating activities", "投资活动": "Investing activities", "资本开支": "Capital expenditure", "投资活动现金流净额": "Net cash from investing activities",
    "筹资活动": "Financing activities", "借款提款": "Loan drawdowns", "筹资活动现金流净额": "Net cash from financing activities", "现金净增加额": "Net change in cash", "期初现金": "Opening cash", "期末现金": "Closing cash", "现金 − 最低现金要求（负数=缺口）": "Cash − minimum cash (negative = gap)",
    "自由现金流（用于 IRR / NPV）": "Free cash flow (for IRR / NPV)", "无杠杆自由现金流": "Unlevered free cash flow", "累计无杠杆自由现金流": "Cumulative unlevered FCF", "期末价值（仅最后一年）": "Terminal value (final year only)", "无杠杆自由现金流 + 期末价值": "Unlevered FCF + terminal value",
    "股东现金流（−投入 + 分红 + 期末权益价值）": "Equity cash flow (−injections + dividends + terminal equity value)",
    # 指标
    "关键指标": "Key indicators", "逐年指标": "Indicators by year", "净资产收益率（ROE）": "Return on equity (ROE)", "利息保障倍数（EBIT/利息）": "Interest coverage (EBIT / interest)", "偿债覆盖率 DSCR（(EBITDA−税)/本息）": "DSCR ((EBITDA − tax) / debt service)",
    "经营现金流 / 本息": "Operating cash flow / debt service", "盈亏平衡收入（固定成本/毛利率）": "Break-even revenue (fixed costs / gross margin)", "盈亏平衡收入 / 实际收入": "Break-even revenue / actual revenue", "汇总指标": "Summary indicators",
    "项目 IRR（全投资口径，含期末价值）": "Project IRR (all-investment, with terminal value)", "项目 IRR（全投资口径，不含期末价值）": "Project IRR (all-investment, no terminal value)", "项目 NPV（含期末价值）": "Project NPV (with terminal value)", "股权 IRR（含期末权益价值）": "Equity IRR (with terminal equity value)",
    "最低期末现金": "Minimum closing cash", "最低现金出现年份": "Year of minimum cash", "资金缺口（最低现金要求 − 最低期末现金，≤0 表示无缺口）": "Funding gap (minimum cash − minimum closing cash; ≤ 0 = none)", "股权投入合计": "Total equity injected", "借款提款合计": "Total loan drawdowns",
    "政府补助合计": "Total grants", "资本开支合计": "Total capex", "预测期累计净利润": "Cumulative net income (forecast period)", "预测期累计经营现金流": "Cumulative operating cash flow (forecast period)", "最低 DSCR（有偿债年份）": "Minimum DSCR (years with debt service)", "收入 CAGR（第1年→末年）": "Revenue CAGR (year 1 → final year)",
    # 情景对比
    "情景对比（由 recalc_and_extract.py 写入静态值）": "Scenario comparison (static values written by recalc_and_extract.py)", "运行 scripts/recalc_and_extract.py 后此表将填入三种情景的关键结果。": "Run scripts/recalc_and_extract.py to fill this sheet with the three scenarios' key results.",
    # 月度
    "季节性占比": "Seasonality share", "收款滞后（月）": "Collection lag (months)", "付款滞后（月）": "Payment lag (months)", "现金流入": "Cash inflows", "销售收款": "Collections from sales", "流入合计": "Total inflows", "现金流出": "Cash outflows", "材料/直接成本付款": "Payments for materials / direct costs",
    "费用与人工（含生产人工）": "Opex and labour (incl. production labour)", "偿还本金（年末）": "Principal repaid (year-end)", "所得税（年末）": "Income tax (year-end)", "流出合计": "Total outflows", "净现金流": "Net cash flow", "月份": "Month", "单位": "Unit",
    "与年度表差异（月度 − 年度，口径差异）": "Difference vs annual statement (monthly − annual; basis difference)",
    "说明：收入按季节性占比分摊并按收款滞后月数收款；材料成本按季节性分摊并按付款滞后付款；费用与人工按月平均；利息自提款月起平均计提；本金与所得税年末支付。": "Notes: revenue is spread by seasonality and collected after the collection lag; material cost is spread by seasonality and paid after the payment lag; opex and labour are spread evenly; interest accrues evenly from the drawdown month; principal and tax are paid at year-end.",
    "口径差异：年度表按周转天数计算年末应收/应付/存货，月度表按整月滞后且不含存货占用，因此 12 月末现金与年度表期末现金会有差异（见表末“与年度表差异”行）；月度表用于看现金节奏与最低点，年度表为正式口径。": "Basis difference: the annual statements compute year-end receivables / payables / inventory from turnover days, while this sheet uses whole-month lags and no inventory build, so December cash differs from the annual closing cash (see the last row). Use this sheet for cash rhythm and the low point; the annual statements are the formal basis.",
    # 来源日志
    "假设来源与证据日志": "Assumption sources and evidence log", "编号": "ID", "内容": "Item", "来源 / 依据": "Source / basis", "等级": "Grade", "置信度": "Confidence", "影响": "Impact", "需确认": "Confirm", "模型字段引用的登记册编号": "Register IDs referenced by model fields",
    "高": "High", "中": "Medium", "低": "Low", "是": "Yes", "否": "No", "用户提供": "Provided by user",
    # 说明表
    "结构：假设 → 收入 / 费用 / 投资与折旧 / 融资 → 利润表 → 现金流量表 → 资产负债表 → 指标 → 情景对比。": "Structure: Assumptions → Revenue / Opex / Capex & Depreciation / Financing → Income Statement → Cash Flow → Balance Sheet → Indicators → Scenarios.",
    "颜色：蓝字黄底 = 输入（只改这些）；黑字 = 公式；灰底 = 合计。": "Colours: blue on yellow = inputs (change only these); black = formulas; grey = totals.",
    "情景：在“假设”表的“情景选择”填 1（基准）、2（下行）、3（上行），全模型随之重算；“情景对比”表由脚本写入三情景的关键结果（静态值）。": "Scenarios: type 1 (base), 2 (downside) or 3 (upside) in the scenario selector on the Assumptions sheet and the whole model recalculates; the Scenarios sheet holds static results for all three, written by the script.",
    "检查：资产负债表最后一行“检查差额”应为 0；现金流量表“期末现金”应始终 ≥ 最低现金要求，否则存在资金缺口。": "Checks: the balance check row on the Balance Sheet must be 0; closing cash on the Cash Flow sheet should stay ≥ the minimum cash requirement, otherwise there is a funding gap.",
    "政府补助：按不占股的资本投入处理（计入资本公积与筹资现金流，不计税）；若为收益性补助或需按递延收益分摊，请在正文说明口径。": "Grants are treated as non-equity capital (capital reserve, financing cash flow, untaxed); if a grant is revenue-related or must be deferred and released, state the basis in the plan.",
    "简化处理：年度模型，期末计息、等额本金还款；所得税按利润总额×税率，亏损可结转抵扣（未设 5 年期限）；不含增值税与附加税；折旧直线法；营运资金按周转天数估算。": "Simplifications: annual model, year-end interest, equal-principal repayment; income tax = profit before tax × rate with unlimited loss carry-forward; no VAT; straight-line depreciation; working capital from turnover days.",
    "免责：本模型基于假设表中的假设，用于商业计划分析，不构成审计意见、税务或法律建议，也不构成收益或融资承诺。": "Disclaimer: this model rests on the assumptions in the Assumptions sheet and supports business-plan analysis; it is not an audit opinion, tax or legal advice, nor a promise of returns or financing.",
    "由 Bankable skill 的 build_financial_model.py 生成。": "Generated by build_financial_model.py from the Bankable skill.",
}

# 模式规则：(正则, 替换)，按顺序应用于词典未命中的字符串
PATTERNS = [
    (re.compile(r"^折旧：(.+)$"), r"Depreciation: \1"),
    (re.compile(r"^在建：(.+)$"), r"CIP: \1"),
    (re.compile(r"^资本开支：(.+)$"), r"Capex: \1"),
    (re.compile(r"^借款：(.+)$"), r"Loan: \1"),
    (re.compile(r"^费用：(.+)$"), r"Opex: \1"),
    (re.compile(r"^收入线：(.+)（销量/单价/成本）$"), r"Revenue line: \1 (volume / price / cost)"),
    (re.compile(r"^(.+) — 人数（(.+)）$"), lambda m: f"{m.group(1)} — FTE ({CATEGORIES.get(m.group(2), m.group(2))})"),
    (re.compile(r"^(.+) — 人均年成本$"), r"\1 — annual cost per FTE"),
    (re.compile(r"^   (.+) — 还本计划（按年）$"), r"   \1 — repayment schedule (by year)"),
    (re.compile(r"^   (.+) — 提款计划（按年）$"), r"   \1 — drawdown schedule (by year)"),
    (re.compile(r"^   依据：(.+)$"), r"   Basis: \1"),
    (re.compile(r"^(.+)（依据：(.+)）$"), r"\1 (basis: \2)"),
    (re.compile(r"^第 (\d+) 年（(\d{4})）月度现金流（参考口径）$"), r"Year \1 (\2) monthly cash flow (indicative)"),
    (re.compile(r"^年度表期末现金（第 (\d+) 年）$"), r"Closing cash per annual statement (year \1)"),
    (re.compile(r"^月度现金流\(第(\d+)年\)$"), r"Monthly Cash (Year \1)"),
    (re.compile(r"^(.+) — 财务模型说明$"), r"\1 — financial model notes"),
    (re.compile(r"^生成日期：(\S+)\s+货币单位：(\S+)（除注明外）\s+预测期：(\S+)（(\d+) 年）$"), r"Generated: \1    Unit: \2 (unless stated)    Forecast period: \3 (\4 years)"),
    (re.compile(r"^情景依据：(.+)$"), r"Scenario rationale: \1"),
    (re.compile(r"^下行：(.+)$"), r"Downside: \1"),
    (re.compile(r"^上行：(.+)$"), r"Upside: \1"),
    (re.compile(r"^(.+)（(销售费用|管理费用|研发费用|其他费用|营业成本)）$"), lambda m: f"{m.group(1)} ({CATEGORIES[m.group(2)]})"),
]
# 单位串里的中文
UNIT_TOKENS = [("万元/人", "{unit}/FTE"), ("万元", "{unit}"), ("/吨", "/t"), ("吨", "t"), ("/人", "/FTE"), ("人", "FTE"), ("天", "days"), ("年", "yrs"), ("月", "mo")]

CJK = re.compile(r"[一-鿿]")


def translate_label(text, unit="万元"):
    if not isinstance(text, str) or text.startswith("=") or not CJK.search(text):
        return text
    if text in EXACT:
        return EXACT[text]
    for pat, rep in PATTERNS:
        if pat.search(text):
            out = pat.sub(rep, text)
            # 替换结果里可能还有类别词或其他固定短语
            for k, v in CATEGORIES.items():
                out = out.replace(f"（{k}）", f" ({v})")
            return out
    # 单位串
    if len(text) <= 6:
        for k, v in UNIT_TOKENS:
            if text == k:
                return v.format(unit=unit)
        t = text
        for k, v in UNIT_TOKENS:
            t = t.replace(k, v.format(unit=unit))
        if not CJK.search(t):
            return t
    # 类别词嵌在其他字符串里
    t = text
    for k, v in CATEGORIES.items():
        t = t.replace(f"（{k}）", f" ({v})").replace(k, v)
    for k, v in EXACT.items():
        if len(k) >= 2 and k in t and not k.isascii():
            t = t.replace(k, v)
    return t


def rename_sheets_and_formulas(wb, monthly_sheet=None):
    """重命名工作表并改写所有公式里的表引用（含带引号与不带引号两种写法）。"""
    mapping = dict(SHEET_NAMES)
    for ws in wb.worksheets:
        m = re.match(r"^月度现金流\(第(\d+)年\)$", ws.title)
        if m:
            mapping[ws.title] = f"Monthly Cash (Year {m.group(1)})"
    # 按长度降序替换，避免“收入”误伤“收入线”之类（表名里不含，但稳妥）
    pairs = sorted(mapping.items(), key=lambda kv: -len(kv[0]))
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    for zh, en in pairs:
                        v = v.replace(f"'{zh}'!", f"'{en}'!").replace(f"{zh}!", f"'{en}'!")
                    c.value = v
    for ws in wb.worksheets:
        if ws.title in mapping:
            ws.title = mapping[ws.title]
    # _map 表的 sheet 列
    if "_map" in wb.sheetnames:
        ws = wb["_map"]
        for row in ws.iter_rows(min_row=2):
            if row[1].value in mapping:
                row[1].value = mapping[row[1].value]
    return mapping


def localize_workbook(wb, unit="万元"):
    """把整个工作簿本地化为英文；返回仍含中文的单元格数量（用户内容除外无法判断，供日志参考）。"""
    rename_sheets_and_formulas(wb)
    remaining = 0
    for ws in wb.worksheets:
        if ws.title == "_map":
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith("=") and CJK.search(v):
                    nv = translate_label(v, unit)
                    c.value = nv
                    if CJK.search(nv):
                        remaining += 1
    return remaining


SCENARIO_SHEET_EN = {
    "指标": "Metric", "基准": "Base", "下行": "Downside", "上行": "Upside", "单位": "Unit",
    "营业收入（{y}）": "Revenue ({y})", "收入 CAGR": "Revenue CAGR", "毛利率（{y}）": "Gross margin ({y})", "EBITDA（{y}）": "EBITDA ({y})", "净利润（{y}）": "Net income ({y})",
    "预测期累计净利润": "Cumulative net income", "首个盈利年份": "First profitable year", "首个经营现金流为正年份": "First year of positive operating cash flow", "最低期末现金": "Minimum closing cash",
    "最低现金出现年份": "Year of minimum cash", "资金缺口（相对最低现金要求）": "Funding gap (vs minimum cash)", "峰值资金需求（无融资口径）": "Peak funding need (before financing)", "最低 DSCR（全部年份）": "Minimum DSCR (all years)",
    "最低 DSCR（经营期，有还本年份）": "Minimum DSCR (operating years)", "项目 IRR（含期末价值）": "Project IRR (with terminal value)", "项目 IRR（不含期末价值）": "Project IRR (no terminal value)", "项目 NPV": "Project NPV", "回收期（无杠杆 FCF，年）": "Payback (unlevered FCF, years)",
    "营业收入": "Revenue", "净利润": "Net income", "期末现金": "Closing cash", "经营活动现金流": "Operating cash flow", "（逐年，{u}）": " (by year, {u})",
    "三情景关键结果（静态值，由 recalc_and_extract.py 于 {d} 写入；修改假设后需重新运行脚本）": "Key results for the three scenarios (static values written by recalc_and_extract.py on {d}; rerun the script after changing assumptions)",
    "年": "", "x": "x",
}
