# 财务模型指南

目录
- [模型思路](#模型思路)
- [工作簿结构](#工作簿结构)
- [假设文件 assumptions.json 字段说明](#假设文件-assumptionsjson-字段说明)
- [填假设的方法](#填假设的方法)
- [情景设计](#情景设计)
- [模型输出 model_outputs.json](#模型输出-model_outputsjson)
- [检查项与常见修正](#检查项与常见修正)
- [各类读者看哪些指标](#各类读者看哪些指标)
- [把数字写进正文](#把数字写进正文)
- [模型的简化与局限（交付时要说明）](#模型的简化与局限交付时要说明)

## 模型思路

模型回答四个问题：这门生意的利润结构是什么（利润表）、钱什么时候进出（现金流量表）、需要多少外部资金且什么时候需要（资金缺口与峰值需求）、关键假设变化时结论还成不成立（情景）。

驱动链：销量 × 单价 → 收入；销量 × 单位成本（或收入 × 成本率）+ 生产人工 → 营业成本；费用明细 + 人员 → 费用；资本开支 → 折旧与固定资产；周转天数 → 营运资金；股权与借款 → 融资现金流与利息。三张表互相勾稽，资产负债表自动平衡，不允许"插数"。

年度模型（3–10 年），另有第 1 年月度现金流作为参考口径。所有金额单位统一（默认万元）。

## 工作簿结构

| 工作表 | 内容 | 谁改 |
|---|---|---|
| 说明 | 结构、颜色约定、简化处理、免责 | 不改 |
| 假设 | 全部输入（蓝字黄底）：情景选择与乘数、收入/费用/人员/资本开支/营运资金/融资/税务与估值/期初余额 | 用户和你 |
| 收入 | 情景后的销量、单价、收入、营业成本、毛利 | 公式 |
| 费用 | 费用明细、人员成本、分类合计（销售/管理/研发/其他）、生产人工 | 公式 |
| 投资与折旧 | 资本开支、直线折旧、固定资产净值 | 公式 |
| 融资 | 股权投入、每笔借款的提款/还本/利息/余额、存量借款 | 公式 |
| 利润表 | 收入→毛利→EBITDA→EBIT→利润总额→所得税（亏损结转）→净利润→分红→留存 | 公式 |
| 现金流量表 | 间接法三段 + 期末现金 + 无杠杆自由现金流 + 期末价值 + 股东现金流 | 公式 |
| 资产负债表 | 货币资金/应收/存货/其他流动资产/固定资产；应付/借款；实收资本/未分配利润；检查差额 | 公式 |
| 指标 | 逐年比率（毛利率、EBITDA 率、净利率、ROE、资产负债率、利息保障倍数、DSCR、盈亏平衡）+ 汇总（IRR、NPV、最低现金、资金缺口、累计值、CAGR） | 公式 |
| 情景对比 | 三情景关键结果（静态值，脚本写入） | 脚本 |
| 月度现金流(第1年) | 季节性收款、付款滞后、投资与融资时点、月末现金 | 公式 |
| 来源日志 | 假设的登记册编号与来源 | 你 |
| _map（隐藏） | 行位置索引，供脚本读取 | 不改 |

## 假设文件 assumptions.json 字段说明

示例见 `assets/assumptions-example.json`。时间序列字段三种写法：数字（每年相同）、数组（长度 = 预测年数）、`{"year1": 值, "growth": 增长率}`（growth 可为单个数或长度为 years−1 的数组）。growth 写法在 Excel 中保留为"首年输入 + 逐年增长率"两行，用户可直接改。

```
meta
  project_name        项目名称（出现在模型与输出中）
  currency            币种，仅标注
  unit                金额单位，默认 "万元"；全部金额按此单位
  start_year          首个预测年，如 2027
  years               预测年数 3–10
  existing_business   true 时必须提供 opening_balance

scenarios                        见"情景设计"
  downside / upside: {volume, price, unit_cost, opex, label}   乘数（1 = 不变）
  rationale                      情景依据（写风险编号）

revenue_lines[]                  每条收入线
  id, name, unit                 unit 如 "吨""台""项目""客户"，为空表示按金额
  volume                         销量序列
  price                          单价序列（单位 = unit/unit 产品）
  unit_cost 或 cogs_pct          二选一：单位成本序列，或成本率序列（占收入）
  source                         登记册编号

opex_lines[]                     固定/半固定费用
  id, name, category             category ∈ 销售费用/管理费用/研发费用/其他费用
  values                         序列，或 {"pct_of_revenue": 0.04} 按收入比例
  source

headcount[]                      人员（成本 = 人数 × 人均年成本）
  role, category                 category 可为上述四类或 "营业成本"（生产人工计入营业成本）
  count                          人数序列
  annual_cost, cost_growth       首年人均年成本、年涨幅

capex[]                          资本开支
  name, amount, year             year = 第几个预测年（1 起）
  life_years, salvage_pct        折旧年限、残值率（默认 0）

working_capital
  ar_days, inventory_days, ap_days, other_current_assets_pct
  （应收按收入、存货按营业成本、应付按营业成本+费用）

financing
  equity[]: {year, amount, note}
  grants[]: {year, amount, note}  政府补助 / 不占股资本投入：计入筹资现金流与资本公积，不计税
            （收益性补助或按递延收益分摊的，在正文说明口径；若想当作收入，改放 revenue_lines）
  loans[]:  {id, name, amount, year, rate, term_years, grace_years}
            等额本金；宽限期内只付息；还本年份 = 提款年+宽限期 … 提款年+期限−1
  dividend_payout                分红比例（占正净利润），默认 0
  use_of_funds[]                 {name, amount}，用于资金用途图与融资章节，不参与计算

tax
  income_tax_rate                默认 0.25；亏损可结转抵扣（无年限限制，见局限）
  opening_loss_carryforward      期初可抵扣亏损（正数填，脚本记为负池）

valuation
  discount_rate                  NPV 折现率（建议 10–15%，按行业与阶段）
  terminal_method                multiple（EBITDA 倍数）/ gordon（永续增长）/ none
  ebitda_multiple, terminal_growth

min_cash                         最低现金要求；低于此值视为资金缺口

monthly_year1（可选）
  seasonality[12]                各月收入占比，和为 1
  collection_lag_months, payment_lag_months   省略时按 ar_days/30、ap_days/30 取整推出，与年度口径一致
  equity_month, loan_month, grant_month        资金到位月份
  capex_months: {资本开支 id: 月份}

opening_balance（存量业务，或已有少量经营的初创企业）
  cash, accounts_receivable, inventory, other_current_assets, fixed_assets_net, fixed_assets_remaining_life,
  accounts_payable, loans, loan_rate, loan_remaining_years, paid_in_capital, capital_reserve, retained_earnings
  （资产必须 = 负债 + 权益，否则脚本报错；通常用 retained_earnings 配平，并在登记册注明）

sources[]（可选）                 {id, content, source, grade, confidence, impact, confirm}，写入来源日志
```

脚本会校验类型、长度、范围，错误信息指向具体字段（如 `revenue_lines[0].volume.growth: 增长率数组长度应为 4`）。

## 填假设的方法

优先级：用户数据（合同、报价、历史报表）> 行业基准（可比公司年报、协会数据）> 推断（写明推理）。每个字段的 `source` 写登记册编号，这样来源日志和登记册能对上。

- **销量**：必须受能力约束。写出"产能 × 利用率"或"客户数 × 转化率 × 单客采购量"的推导，首年利用率/转化率用行业基准。
- **单价**：有报价或可比价格区间就取中值，并在登记册写区间。
- **单位成本/成本率**：BOM、采购报价或可比公司毛利率反推；注意是否含运费、包装、损耗。
- **费用**：租金、研发、推广按合同或预算；推广也可按收入比例。人员用岗位表，人均成本含社保。
- **资本开支**：设备报价、工程概算；分期投入按进度表。
- **营运资金**：客户账期、供应商账期、库存周期——这是初创企业现金流被低估的首要原因。
- **融资**：股权金额与到位时间；贷款的金额、期限、利率、宽限期按银行意向。
- **税**：企业所得税率；享受优惠要有政策依据。增值税不在模型内。
- **折现率与期末价值**：早期项目折现率高（15–25%），成熟项目 8–12%；EBITDA 倍数参考可比交易；无依据时用 none 并说明。

填完先跑一遍，看 `checks` 和 `warnings`，再调整，不要一次就定稿。

**历史期怎么衔接**：模型只从 `start_year` 起算，没有"历史期"列。
- 成熟企业：`existing_business: true`，期初资产负债表取最近一期报表，历史收入与利润放正文 11.9 节的表格（来自用户报表，不进模型）。
- 已有少量收入的初创企业（例如今年签了 180 万元合同）：同样用期初余额——把截至 `start_year` 前的现金、应收、已签未执行合同对应的预收/应收、实收资本填入 `opening_balance`，用 retained_earnings 配平；当年尚未完成的合同收入放进第 1 年的销量假设里，并在登记册写明"含上年结转合同 ×× 万元"。不要把历史收入塞进第 1 年去凑数。
- 纯构想期：`existing_business: false`，期初全为 0，第 1 年即建设/起步年。

**第 1 年月度表与年度表的差异**：月度表按整月收付款滞后分摊、不含存货占用，年度表按周转天数计算年末应收/应付/存货，所以 12 月末现金与年度表期末现金通常不同；`model_outputs.json` 的 `monthly_year1.difference_vs_annual` 给出差额。正文统一写法："月度表用于判断现金节奏与最低点（最低点出现在 × 月，约 ×× 万元），正式口径以年度表为准。"

## 情景设计

情景乘数作用于**所有收入线**的销量、单价、单位成本（或成本率）和所有费用：

- 下行：从风险矩阵里挑概率 ≥ 中的风险，量化成乘数（如渠道延迟 → 销量 ×0.85；价格战 → 单价 ×0.97；原料涨价 → 单位成本 ×1.03），并在 `label` 里写明对应风险编号。不要用"统一打八折"这种无依据的设定。
- 上行：只放有机制支撑的利好（已签意向、政策落地、规模采购）。
- 读者偏好：银行版正文重点展示下行；投资人版展示三情景；政府版通常只需基准 + 下行。

下行情景出现资金缺口是常见且有价值的发现——它决定了"应对措施"章节要写什么（备用授信、分期投入、费用弹性、里程碑式融资）。

**单因素敏感性**：`chapter-guide` 要求的"关键假设 ±20% 敏感性"用脚本做，不要手工复制假设文件：

```bash
python scripts/run_sensitivity.py work/assumptions.json -o work/sensitivity_outputs.json \
  --md work/sensitivity.md --chart work/charts/sensitivity.png [--cases my_cases.json]
```

默认 12 个案例（销量 ±20%、价格 ±10%、单位成本 ±10%、费用 ±10%、资本开支 ±20%、应收账期 ±30 天），输出表格（可直接贴进 11.8 节）、龙卷风图和 JSON。正文引用这些数字时，把 `sensitivity_outputs.json` 作为 `check_consistency.py --extra` 传入。要测"没有 B 轮""二期推迟"这类结构性备选方案，用 `--cases` 写 `patch: json` 案例，或复制假设文件单独跑。

## 模型输出 model_outputs.json

```
generated, model_file, project_name, unit, years[], min_cash_requirement
scenarios.{base|downside|upside}
  label
  series.*            逐年数组：revenue, cogs, gross_profit, gross_margin, revenue_growth, opex_total, opex_各类,
                      headcount_total, capex, depreciation, equity_in, loan_drawdown, loan_repayment, interest,
                      debt_balance, debt_service, ebitda, ebitda_margin, ebit, ebt, tax, net_income, net_margin,
                      cfo, cfi, cff, closing_cash, cash_vs_min, unlevered_fcf, cumulative_ufcf, terminal_value,
                      total_assets, total_liabilities, total_equity, balance_check, debt_ratio, roe,
                      interest_coverage, dscr, breakeven_revenue, breakeven_ratio …
  summary.*           irr_with_tv, irr_no_tv, npv, equity_irr, min_cash_balance, min_cash_year, funding_gap,
                      total_equity_in, total_grant_in, total_loan_drawdown, total_capex, total_net_income, total_cfo, dscr_min, revenue_cagr
  derived.*           first_profit_year, first_positive_cfo_year, payback_years_ufcf（超出预测期为 null）,
                      peak_funding_need（无融资口径的峰值需求）, negative_cash_years, cash_below_min_years,
                      dscr_by_year, dscr_below_1_2_years, gross_margin_range, max_revenue_growth
  monthly_year1       closing_cash[12], net_cash[12], min_closing_cash, min_month, annual_year1_closing_cash,
                      difference_vs_annual, explanation（有月度表时）
checks                balance_sheet_balances, funding_gap_*, dscr_*, warnings[]
narrative_numbers     已格式化的字符串（"2,120.0 万元""39.4%"），正文直接引用
use_of_funds, scenario_labels（传入 --assumptions 时）
```

## 检查项与常见修正

| 现象 | 通常原因 | 修正 |
|---|---|---|
| `balance_sheet_balances: false` | 期初余额不平；手改了公式 | 用 retained_earnings 配平期初；重新生成 |
| 基准情景现金为负 / 低于最低现金 | 投资前置、营运资金被低估、融资到位晚 | 增加或提前融资；分期投入；缩短账期；降低首年费用 |
| 下行情景资金缺口 | 利润薄、固定成本高 | 风险章节写应对：备用授信、费用弹性、里程碑式投入 |
| DSCR < 1.2 | 还款期太短、宽限期不足、前期利润薄 | 延长期限或宽限期、调整还款节奏、增加股权占比 |
| 毛利率 > 85% 或 < 5% | 单价/成本口径错（含税/不含税、是否含运费） | 核对假设口径 |
| 某年增长 > 100% | 没有产能/获客约束；小基数（脚本会区分） | 大基数：在正文写明产能与获客机制，或下调；小基数：一句话说明爬坡即可 |
| 下行情景 DSCR 为负 | 下行情景 EBITDA 为负 | 正常现象，正文写"下行情景下第 × 年无法以经营现金流覆盖本息，需 ×× 万元备用资金"而不是只贴数字 |
| IRR 为空（"-"） | 现金流全为负或无符号变化 | 检查期末价值设置；项目可能确实不可行 |
| 回收期 null | 预测期内未回收 | 正文写"超出预测期"，考虑延长预测年数或说明长期逻辑 |

## 各类读者看哪些指标

| 读者 | 正文重点展示 |
|---|---|
| 投资人 | 收入与 CAGR、毛利率走势、EBITDA 转正年、净利润、单位经济（正文补充）、资金用途、峰值资金需求、项目/股权 IRR（含期末价值，说明倍数依据）、三情景 |
| 银行 | DSCR 逐年与最低值、利息保障倍数、资产负债率、经营现金流/本息、最低现金、下行情景偿债能力、月度现金流 |
| 政府 | 总投资、资金筹措（含配套）、达产年收入/利润/税收、投资回收期、IRR（可研口径常用不含期末价值）、就业人数（人员总数） |
| 内部 | NPV、IRR、回收期、峰值资金需求、敏感性、对公司现金的影响 |

## 把数字写进正文

- 只用 `narrative_numbers` 或 `series/summary` 里的值；四舍五入到整数万元或一位小数，全篇统一。
- 写口径："基准情景下""含期末价值""不含增值税"。
- 百分比一位小数；倍数两位小数；年份写公元年。
- 表格里的年度数字与 `series` 一一对应，不要在表格里手算合计。
- 情景结论写成对比句："下行情景下第五年收入 4,692 万元（较基准低 18%），最低现金 1 万元，需 149 万元备用资金。"

## 模型的简化与局限（交付时要说明）

- 年度模型，期末计息、期末收付；月度表仅为参考口径。
- 等额本金还款；不支持等额本息、气球贷（可在 Excel 里手改还本行）。
- 所得税 = 应纳税所得 × 税率，亏损无限期结转（实际多为 5 年）；不含增值税、附加税、递延税。
- 折旧直线法；无在建工程、无无形资产摊销（可将研发资本化项目作为 capex 处理）。
- 营运资金按周转天数线性估算。
- 期末价值用 EBITDA 倍数或永续增长，仅用于 IRR/NPV 参考，不是估值结论。
- 未经审计；不构成投资、贷款或税务建议。
