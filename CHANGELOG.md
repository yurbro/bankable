# Changelog

## v0.4.0 — 2026-11-02

English-first engagements, from the UK dairy-farm (GBP) and Rotterdam rPET (EUR) eval runs:

- `build_financial_model.py --lang en`: English sheet names and labels (formula references rewritten, user-written names kept); `recalc_and_extract.py --lang en` for the scenario sheet, notes and `narrative_numbers` (£k/€k/$k formatted as "£2,040k" / "£2.87m"); `model_i18n.py` holds the dictionary. English category aliases (selling, g&a/overheads, r&d, other, cogs/direct labour).
- Model: `opening_balance.land` (non-depreciating), `financing.fixed_distributions` (partners' drawings), loan `schedule` may sum to less than the loan when the term runs past the forecast horizon.
- Existing businesses: incremental initial working capital, no spurious total-investment warning, `use_of_funds` shortfall reported as "to be funded from operating cash flow".
- `check_consistency.py --lang en` (English report), Unicode minus and sign-insensitive matching, currency-prefixed amounts (£2.2m, GBP 2,200k, €4 million), thousand-unit models.
- `run_sensitivity.py`: `--cases` appends to the defaults (`--cases-only` to replace), `patches` for multi-path cases, `--variants path@downside`, ASCII brackets in English, one-sided cases coloured by sign, equity IRR.
- `make_tables.py`: all-zero rows and empty IRR rows dropped, `--category-labels` (e.g. Overheads for UK), incremental working-capital label for existing businesses.
- `build_docx.py`: content-weighted column widths, English placeholders.
- English register and review-memo templates (`assets/*.en.md`); bilingual guide gains an "English-first engagements" section (GBP/EUR conventions, UK/EU/US terms); SKILL.md decides the deliverable language at step 0.
- README now English by default with a language switch; English diagram set; UK and EU eval cases and sample runs added.

## v0.3.0 — 2026-11-02

根据银行贷款（存量合作社）和政府申报 + 董事会（双语、建设期项目）两次试跑的反馈：

- 模型：情景乘数支持逐年数组（周期路径）；贷款支持 `schedule`（任意/到期一次还本）与 `drawdown`（分期提款）；存量借款支持 `loan_schedule`；资本开支 `in_service_year`（投用前计入在建工程，投用年起折旧）；月度现金流可选任意一年（`year_index`）；流动比率分母加入一年内到期借款；资产负债表新增在建工程行。
- 输出：导出全部序列（分产品线收入/成本/销量、情景乘数）；经营期 DSCR（有还本年份）；建设期年份、建设期利息、铺底流动资金、可研口径总投资与资金来源对账；首个经营年起算的 CAGR；无前期净投入或离谱时 IRR 置空并给 `irr_note`；`use_of_funds` 与资金来源不一致、资金来源与总投资差异超 15% 时警告；毛利率与增长率警告忽略无收入年份。
- 新脚本 `make_tables.py`：从模型输出生成正文全部财务表格（中/英，万元或 RMB million），杜绝手工转录。
- `run_sensitivity.py`：`--variants` 把整套备选方案列入同一张表；新增股权 IRR；`--lang en`；json patch 支持列表索引（`financing.loans.0.rate`）。
- `check_consistency.py`：首年无收入时以首个经营年为必提数字；经营期 DSCR；识别银行/政府/董事会章名；亿元写法须紧跟"亿"、数字带边界匹配；识别英文 "RMB x million"；`--extra` 文件中按年数组不再丢弃，比率/百分比分类更准。
- `build_pptx.py`：字号随要点数量自动调整、长标题缩小、Markdown 表格转 PPT 表格、图按比例适配；`make_charts.py` 无收入年份不标数字。
- 指南：start_year 选法、周期情景、建设期与可研口径、DSCR 两个口径、补助的递延收益说明与"无补助"敏感性、同时两个决策时主读者的选法与双执行摘要示例、英文版流程、PPT 每页要点上限。

## v0.2.0 — 2026-11-02

根据第一次完整试跑（投资人版、无输入数据）的反馈改进：

- 新增 `scripts/run_sensitivity.py`：单因素敏感性（销量/价格/单位成本/费用/资本开支/应收账期），输出表格、龙卷风图与 JSON；支持自定义案例。
- 第 1 年月度现金流：收付款滞后默认由周转天数推出；表末增加"与年度表差异"行；`model_outputs.json` 给出差额与标准解释措辞。
- 收入增长 >100% 的警告区分"小基数爬坡"与"大基数"。
- 财务模型支持政府补助 / 不占股资本投入（`financing.grants`，计入资本公积与筹资现金流）；期初余额支持 `capital_reserve`。
- `check_consistency.py`：紧跟 `[A-xx]`/`[S-xx]` 等标记的数字自动跳过；纳入差额、占比、累计、情景差异、资金用途占比等派生值；`--extra` 可传入敏感性输出。误报明显减少。
- `build_docx.py --strip-refs`：对外终稿一键去除登记册/来源标记（`[待补充]` 保留）；封面不显示页眉页脚；目录域预填标题列表。
- SKILL.md：新增"精简版"模式；说明模型与章节写作的迭代关系；交付前去标记与脱敏提醒。
- 指南：有少量收入的初创企业如何设期初余额；月度/年度口径差异的写法；评审备忘编号改为 Q-xx。

## v0.1.0 — 2026-11-01

首版：SKILL.md 八步流程；9 份参考文件；财务模型 / 重算与三情景 / 图表 / Word / PPT / 一致性核对脚本；假设示例、章节骨架、登记册与评审备忘模板；3 个测试用例。
