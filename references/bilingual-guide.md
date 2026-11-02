# 英文版 / 双语规范

英文版给境外投资人、合作伙伴或国际机构看，目标是让他们在不了解中国语境的情况下读懂并信任。做法是**改写**，不是逐句翻译。

## 流程

1. 中文版定稿并通过核对后，再做英文版（`work/plan_en.md`），结构、章节号、表格、图表与中文版一致，数字一字不差。
2. 英文表格与图：`make_tables.py model_outputs.json -o tables_en.md --lang en`（自动换算 RMB million）、`make_charts.py --lang en -o charts_en/`、`run_sensitivity.py --lang en --chart charts_en/sensitivity.png`。
3. `scripts/build_docx.py plan_en.md --lang en --title "... Business Plan" ...` 生成英文 Word。
4. 跑 `check_consistency.py plan_en.md model_outputs.json --extra sensitivity_outputs.json`：脚本识别 "RMB 31.3 million"、"RMB 10k" 表格和百分比写法；million 写法保留一到两位小数（313.2 million）才能匹配。
5. 登记册和 Excel 模型默认保留中文（登记册是给用户确认的，模型表头可在 Excel 里自行翻译）；需要英文版登记册时只翻译"内容"和"来源"两列的摘要。英文 PPT：按 `pitch-deck-guide` 写 `deck_en.md`，`build_pptx.py --lang en`，图用 `charts_en/`。
6. 交付说明里注明英文版是中文版的对应版本，以中文版为准（或反之，按用户要求）。

## 改写原则

- 段落以结论开头，句子短，主动语态；删掉中文里的套话和政策口号，把其中的实质内容改成事实陈述。
- 中国特有的概念第一次出现时给英文 + 拼音/中文 + 一句解释，例如 *"specialized and sophisticated SMEs that produce new and unique products" ("专精特新" SMEs, a Ministry of Industry and Information Technology designation that confers tax and procurement preferences)*。
- 政策文件：保留中文名称，加英文译名与发布机构、年份、文号，例如 *Action Plan for Reducing Chemical Fertilizer Use (《化肥减量增效行动方案》, MARA, 2025)*。
- 机构名用官方英文名（MARA、NDRC、MIIT、MEE、PBoC、CSRC）。
- 地名用汉语拼音官方拼写；省市加说明（Jiangsu Province, eastern China）。
- 不要把"领导重视""政府支持"写成优势，除非有具体文件、资金或合同。
- 文化上避免过度谦辞和过度宏大叙事，英文读者更在意数据与机制。

## 数字与单位

- 货币：以人民币为准，写 **RMB 21.2 million**；同时在首次出现或关键处给美元换算并注明汇率与日期：*RMB 21.2 million (USD 2.9 million at 7.2 RMB/USD, Nov 2026)*。
- 万元换算：万元 ÷ 100 = 百万元（million）；2,120 万元 = RMB 21.2 million；1.5 亿元 = RMB 150 million。表格可保留 "RMB 10k" 作为单位，但正文优先 million。
- 千分位逗号、小数点英文习惯；百分比一位小数；年份不加"年"。
- 面积、产量用公制并在首次出现时换算亩（1 mu ≈ 0.0667 ha）。
- 财务术语按下表；口径说明（含税/不含税）翻成 *excl. VAT / incl. VAT*。

## 术语表

| 中文 | 英文 |
|---|---|
| 商业计划书 | Business Plan |
| 执行摘要 | Executive Summary |
| 可行性研究报告 | Feasibility Study |
| 项目申报书 | Project Application / Grant Application |
| 公司概况 | Company Overview |
| 股权结构 | Shareholding Structure / Cap Table |
| 实缴资本 / 注册资本 | Paid-in Capital / Registered Capital |
| 市场规模 TAM/SAM/SOM | Total / Serviceable / Obtainable Market |
| 目标客户 | Target Customers |
| 商业模式 | Business Model |
| 单位经济 | Unit Economics |
| 获客成本 / 客户终身价值 | Customer Acquisition Cost (CAC) / Lifetime Value (LTV) |
| 毛利率 | Gross Margin |
| 营业收入 / 营业成本 | Revenue / Cost of Goods Sold (COGS) |
| 销售费用 / 管理费用 / 研发费用 | Selling Expenses / General & Administrative Expenses / R&D Expenses |
| 息税折旧摊销前利润 | EBITDA |
| 利润总额 / 净利润 | Profit Before Tax / Net Income |
| 利润表 / 资产负债表 / 现金流量表 | Income Statement / Balance Sheet / Cash Flow Statement |
| 经营活动现金流 | Operating Cash Flow |
| 营运资金 | Working Capital |
| 应收账款 / 应付账款 / 存货 | Accounts Receivable / Accounts Payable / Inventory |
| 资本性支出 | Capital Expenditure (CapEx) |
| 折旧与摊销 | Depreciation & Amortization |
| 偿债覆盖率 | Debt Service Coverage Ratio (DSCR) |
| 利息保障倍数 | Interest Coverage Ratio |
| 资产负债率 | Debt-to-Asset Ratio |
| 内部收益率 / 净现值 / 投资回收期 | IRR / NPV / Payback Period |
| 盈亏平衡 | Break-even |
| 情景分析 / 敏感性分析 | Scenario Analysis / Sensitivity Analysis |
| 基准 / 下行 / 上行情景 | Base / Downside / Upside Case |
| 资金用途 | Use of Funds |
| 融资金额 / 估值 / 出让比例 | Funding Amount / Valuation / Equity Offered |
| 退出 | Exit |
| 里程碑 | Milestones |
| 风险矩阵 | Risk Matrix |
| 抵押 / 担保 / 质押 | Mortgage (Collateral) / Guarantee / Pledge |
| 宽限期 | Grace Period |
| 等额本金 | Equal Principal Repayment |
| 配套资金 | Matching Funds / Co-financing |
| 专项资金 | Special-purpose Government Fund |
| 绩效目标 | Performance Targets |
| 产业园区 | Industrial Park |
| 高新技术企业 | High-tech Enterprise (HNTE, a tax-status designation) |
| 专精特新 | "Specialized, Refined, Distinctive, Innovative" SMEs |
| 乡村振兴 | Rural Revitalization (national policy program) |
| 碳达峰碳中和 / 双碳 | Carbon Peaking and Carbon Neutrality ("dual carbon" goals) |
| 环评 / 能评 / 安评 | Environmental / Energy-use / Safety Impact Assessment |
| 社会效益 | Social Benefits |
| 保密声明 / 免责声明 | Confidentiality Notice / Disclaimer |

## 双语排版

- 两个独立文件（中、英各一份）比一页双栏更易维护，推荐。
- 用户坚持单文件双语时：每章先中文后英文，表格用双语表头（"营业收入 Revenue"），图表只做一套英文标签或中英并列。
- 文件命名：`<Project>-Business-Plan-<YYYY-MM>.docx`。
