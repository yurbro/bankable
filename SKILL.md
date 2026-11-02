---
name: bankable
description: Bankable——撰写经得起投资人、银行、政府评审和董事会拷问的高级商业计划书（BP）的完整方法。流程：一次性访谈 → 事实/假设/证据分开登记 → 按读者切换章节重点 → 先建可重算的 Excel 三表财务模型（含基准/下行/上行情景、DSCR/IRR/NPV/回收期）→ 以模型数字写叙述 → 叙述与数字自动核对 → 红队评审 → 交付中文 Word（可选中英双语）、Excel 模型、PPT 路演版。只要用户提到商业计划书、BP、融资计划书、项目计划书、可行性研究报告、贷款申请材料、政府项目申报书/资金申请、投资人材料、路演 PPT、财务预测或财务模型，或要求把一个项目想法/一堆资料整理成能给投资人、银行、政府或董事会看的文件，即使没有说出"商业计划书"三个字，也要使用本 skill。Use for any business plan, feasibility study, investor memo, loan proposal, grant application narrative or pitch deck, in Chinese or English, across any industry.
---

# Bankable：经得起拷问的商业计划书

## 这个 skill 做什么

把一个业务构想或一堆零散材料，变成一套**数字自洽、每个主张都有来源或假设编号、按读者定制**的商业计划书交付物：

- `<项目名>-商业计划书-<日期>.docx`（中文正式版；需要时另出英文版）
- `<项目名>-财务模型-<日期>.xlsx`（可重算的三表模型，含情景对比与假设表）
- `<项目名>-路演PPT-<日期>.pptx`（需要时）
- `<项目名>-假设与证据登记-<日期>.md`（给用户逐条确认的清单）

它不是模板填空。它是一套作业方法：**先把事实、假设、推断分开，再让财务模型给出结论，叙述只能引用模型算出来的数字**。缺的信息标 `[待补充：…]`，绝不编造公司历史、客户、合同、资质、团队背景或市场数据。

产出是供审阅的专业首稿，不是审计报告、税务或法律意见，也不是融资承诺——交付时要明说。

## 工作流程总览

| 阶段 | 做什么 | 先读哪个文件 |
|---|---|---|
| 0 盘点 | 读完所有素材，确定读者、产出、语言、环境能力 | 本文件 |
| 1 访谈 | 一轮、最多 8 个问题，只问素材里真没有的 | `references/intake-questions.md` |
| 2 登记 | 建立事实/假设/证据登记册 | `references/research-and-evidence.md` |
| 3 研究 | 联网补行业、政策、竞品、基准数据并引用 | 同上 |
| 4 建模 | 填 `assumptions.json` → 生成 Excel → 重算 → 导出 `model_outputs.json` | `references/financial-model-guide.md` |
| 5 起草 | 按读者权重逐章写作，数字只来自模型输出 | `references/audience-profiles.md`、`references/chapter-guide.md` |
| 6 核对 | 跑一致性核对脚本 + 红队评审，修订 | `references/review-checklist.md` |
| 7 交付 | 生成 Word / Excel / PPT / 英文版，写交付说明 | `references/pitch-deck-guide.md`、`references/bilingual-guide.md` |

每个阶段只读当时需要的参考文件，不要一开始把所有文件都加载进上下文。行业专属的驱动因素和指标在 `references/industry-modules.md`，只在确定行业后读对应小节。

路径约定：下文的 `scripts/`、`references/`、`assets/` 都相对于本 skill 目录（SKILL.md 所在目录），运行脚本时用它的绝对路径；工作文件（假设、模型、正文、图表）放在当前工作目录下的 `work/`，交付物也从那里发出。脚本依赖 Python 3.10+、openpyxl、python-docx、python-pptx、matplotlib，重算 Excel 需要 LibreOffice（`soffice`）；缺少时 `pip install` 补齐，或在交付说明中讲清影响。

## 第 0 步：盘点输入，确定交付

先读完用户给的全部材料（文档、表格、旧版 BP、调研、财务数据），再做任何判断。然后确定五件事：

1. **读者是谁**：股权投资人 / 银行或贷款机构 / 政府申报、补贴或园区招商 / 内部战略、董事会或合作伙伴。可以多选，但要定一个**主读者**决定章节权重，其余读者通过定制的执行摘要和附录覆盖（见 `audience-profiles.md` 的"多读者策略"）。
2. **产出什么**：默认中文 Word + Excel 模型；用户要求时加 PPT、英文版。
3. **业务处于什么阶段**：构想期 / 初创期 / 成长期 / 成熟企业新项目。阶段决定历史数据有多少、预测期多长、哪些章节能写实。
4. **深度模式**：默认做完整版（25–45 页、全部章节、联网研究、三情景 + 敏感性，通常要一个小时上下）。用户说"快速""简版""先出个框架""今晚就要"时做精简版：10–15 页、只保留主读者 ★★★ 章节、研究只查市场规模与政策两项、财务只跑基准 + 下行；交付说明里写明是精简版及省略了什么。
5. **环境能做什么**：检查是否能运行 Python（openpyxl、python-docx、python-pptx）、是否有 LibreOffice（`soffice`，用于重算 Excel）、能否联网。缺什么就在交付说明里讲清楚影响（例如无 LibreOffice 时 Excel 在打开时才会计算，核对脚本改用 Python 侧计算）。

决定主读者时如果用户没说，根据材料判断：提到"融资/估值/股权" → 投资人；"贷款/授信/还款" → 银行；"申报/补贴/专项资金/园区" → 政府；"立项/预算/董事会" → 内部。判断不了就在访谈里问，这是最值得问的一个问题。

## 第 1 步：一次性访谈

访谈的目的是拿到**模型和关键章节真正缺的输入**，不是把所有章节的问题都问一遍。规则：

- 一轮提问，最多 8 个，合并成一条消息，按重要性排序，每个问题说明"为什么需要、没有的话我会怎么假设"。
- 素材里已经有的不问；能靠公开研究得到的不问（行业规模、政策），自己去查。
- 优先级最高的信息通常是：读者与用途、融资/申请金额与资金用途、收入模型（卖什么、给谁、单价、量、现有数据）、主要成本结构与投资额、团队、现有进展（客户、合同、资质）、已知的硬约束（时间、合规、股权）。
- 用户不在或不回答时：按最稳妥的方式假设，每条假设记入登记册并标注"需用户确认"，继续推进，不要停在那里等。

问题清单与追问思路见 `references/intake-questions.md`。

## 第 2 步：事实 / 假设 / 证据登记册

写任何章节之前，先建 `<项目名>-假设与证据登记-<日期>.md`，之后所有章节和模型都引用它。每条记录有：编号（F-事实 / A-假设 / E-推断 / T-待补充）、内容、来源或依据、证据等级（A 官方统计与原始文件 / B 行业报告与权威媒体 / C 一般媒体与企业自述 / D 推断与经验值）、置信度、影响的章节与模型字段、需要谁确认。

这份登记册的价值在于：用户只需要逐条确认或纠正，而不用通读 50 页去找哪些是你编的。格式、证据分级标准、检索策略和不编造清单见 `references/research-and-evidence.md`。

## 第 3 步：研究

能联网时，对市场规模、增长率、政策、竞品、价格和成本基准做检索，优先官方统计、政策原文、上市公司年报/招股书、行业协会，其次券商研报与行业媒体。每条引用在正文里标 `[S1]`，附录列来源清单（名称、机构、日期、链接、查阅日期）。

市场规模优先**自下而上**测算（目标客户数 × 渗透率 × 客单价），自上而下的数字只用于交叉验证，并写出计算过程。不能联网时，把需要核实的数字标为 `[待核实]` 并在登记册写明去哪里查。

## 第 4 步：财务模型先于结论

这是和普通"AI 写 BP"最大的区别：**先把模型跑通，再写增长故事**，叙述只能引用模型算出来的数字，增长有多乐观由模型说了算。

流程：

```bash
# 1. 从 assets/assumptions-example.json 复制一份，按登记册填写假设（每条假设带来源编号）
cp assets/assumptions-example.json work/assumptions.json

# 2. 生成带公式的 Excel 三表模型（会先校验假设文件，报错信息指向具体字段）
python scripts/build_financial_model.py work/assumptions.json -o work/模型.xlsx

# 3. 用 LibreOffice 重算、跑三种情景、导出结构化输出，并把情景对比写回 Excel
python scripts/recalc_and_extract.py work/模型.xlsx -o work/model_outputs.json

# 4. 生成财务图表 PNG，供 Word/PPT 引用
python scripts/make_charts.py work/model_outputs.json -o work/charts/

# 5. 单因素敏感性（销量/价格/成本/费用/资本开支/账期），输出表格、龙卷风图和 JSON
python scripts/run_sensitivity.py work/assumptions.json -o work/sensitivity_outputs.json --md work/sensitivity.md --chart work/charts/sensitivity.png
```

实际操作是迭代的：先用粗假设跑一遍模型，看 SOM 与模型收入是否对得上、资金需求与融资额是否匹配，再写市场与竞争章节；章节写作中发现的新约束（产能、账期、合规时间）回头改假设重跑。每次改假设都重新跑 2–5 步，不要手改输出。

跑完后看 `model_outputs.json` 里的 `checks`：资产负债表是否平、现金是否出现负数、资金缺口多大、DSCR 是否低于银行通常要求的 1.2–1.5、毛利率和增长率是否偏离行业常识。有问题先改假设或融资结构，再回头写。存量企业的扩建项目（首个现金流为正）不要报"项目 IRR"——`irr_note` 会提示，改用有无对比的增量现金流或 NPV / 回收期 / DSCR。

模型的结构、每个假设字段的含义、情景如何从风险分析推导、各类读者看哪些指标，见 `references/financial-model-guide.md`。

## 第 5 步：章节起草

按 `references/chapter-guide.md` 的总目录写，按 `references/audience-profiles.md` 决定哪些章节展开、哪些压缩、执行摘要先回答谁最关心的问题。

写作规则（这些决定了计划书读起来像专业顾问写的还是 AI 写的）：

- **每个主张都有依据**：事实标来源 `[S3]`，假设标编号 `[A-05]`，推断说明推理。没有依据的句子要么删掉，要么改成明确的假设。
- **具体胜过气势**："三年内成为行业领先者"不是计划；"第 3 年在华东市场拿到 8% 份额，对应 2,400 家客户，靠 12 人直销团队和 3 家区域代理实现"才是。
- **数字有单位和口径**：万元/亿元、含税/不含税、口径年份、对比基准。文中的财务数字一律来自 `model_outputs.json`，用同样的四舍五入规则（整数万元或一位小数）。
- **不堆形容词，不用空泛的宏大叙事**；用对比、数据和机制说明优势。避免"赋能、抓手、闭环、生态"之类套话，除非读者是习惯这类表达的政府评审并且有实质内容支撑。
- **风险写真的**：评审最反感"无重大风险"。每类风险给概率、影响、应对、触发指标。
- **执行摘要最后写**，1–2 页，主读者 5 分钟内能决定要不要继续看。

章名按读者调整，但保留章号：第 12 章投资人版叫"融资方案"、银行版叫"贷款方案与偿债能力分析"、政府版叫"资金申请与筹措方案"、董事会版叫"决策请求与资源需求"；一致性核对脚本按这些关键词识别财务相关章节。

把正文写成一个 Markdown 文件（`work/plan.md`），用 `#` 一级标题对应章、`##` 对应节，表格用 Markdown 表格，图片用 `![图 1 收入预测](charts/revenue.png)`，分页用 `<!-- pagebreak -->`。这个文件既是 Word 的源，也是英文版和核对脚本的输入。章节骨架见 `assets/plan-skeleton.md`。

## 第 6 步：核对与评审

两步，都不能跳过：

```bash
# 数字核对：正文里的关键财务数字是否与模型输出一致，财务章节里有没有模型里没有的数字
python scripts/check_consistency.py work/plan.md work/model_outputs.json --extra work/sensitivity_outputs.json
```

报告第二部分"需复核"是启发式结果：紧跟 `[A-xx]`/`[S-xx]` 标记的数字已自动跳过，剩下的逐条看——行业数据、估值倍数、备选方案的数字属正常，但凡是"预测"性质的数字必须能在模型或敏感性输出里找到，找不到就改正文。

然后按 `references/review-checklist.md` 以四种身份（投资人、信贷审查、政府评审、CFO/董事会）各提出最刁钻的 5–8 个问题，加上一致性与写作质量清单，写一份评审备忘（问题、严重度、修改方案），逐条修订正文和模型。严重度为"会被直接毙掉"的问题必须解决或在交付说明中坦白。

## 第 7 步：生成交付物

```bash
# 中文 Word（封面、目录域、编号标题、表格样式、页眉页脚、保密声明）
python scripts/build_docx.py work/plan.md -o "work/<项目名>-商业计划书-<日期>.docx" \
  --title "<项目名>商业计划书" --subtitle "<一句话定位>" --org "<公司/团队>" --date "<YYYY年M月>" --lang zh

# 英文版：先按 bilingual-guide 把 plan.md 改写为 plan_en.md（不是逐句直译），再生成
python scripts/build_docx.py work/plan_en.md -o "work/<Project>-Business-Plan-<date>.docx" --lang en ...

# PPT：按 pitch-deck-guide 写 work/deck.md（每页标题、要点、备注、图），
# 优先用环境里的 pptx skill 做高质量版式；没有时用自带脚本
python scripts/build_pptx.py work/deck.md -o "work/<项目名>-路演PPT-<日期>.pptx" --lang zh
```

Excel 模型在第 4 步已经生成，最终版要用 `recalc_and_extract.py` 处理过的文件（带缓存计算值，打开即显示数字）。用 Word 打开 docx 时提示更新目录是正常的，交付说明里提一句。

给用户审阅的版本保留 `[A-xx]`、`[S-xx]` 等标记，方便对照登记册；用户确认后要发给外部读者的终稿，用 `build_docx.py --strip-refs` 去掉这些标记（`[待补充]` 会保留，提醒补齐），并提醒用户脱敏。

文件名用项目名和日期，不要叫 `output.docx`。所有交付物放在同一个目录，连同登记册一起交付。

## 第 8 步：交付说明

交付时用几句话说清：

1. 读者与版本（主读者是谁，其他读者的定制点在哪）。
2. 用户必须确认的假设（登记册里标"需确认"的，列前 5–10 条影响最大的）。
3. `[待补充]` 项清单，以及它们各自影响哪个结论。
4. 模型的关键结论与检查结果（资金需求、最低现金、三情景结论、DSCR/IRR），以及哪些数字对假设最敏感。
5. 局限：年度模型、简化税务处理、未经审计、不构成投资或贷款承诺。

不要复述文档内容，不要说"我已经完成了以下步骤"。

## 质量底线

遇到下面任何一条，停下来修，不要交付：

- 叙述里出现模型里没有的财务数字，或同一指标在两处不一致。
- 任何一条"事实"没有来源，或者来源是你自己推断出来的。
- 公司历史、客户名单、合同、资质、团队履历有任何一项是编的。
- 资产负债表不平、现金为负却没有对应的融资安排、增长率没有机制支撑。
- 风险章节写的是模板风险，没有针对这个项目的触发条件和应对。
- 执行摘要回答不了主读者的第一个问题（投资人：为什么这是个好生意；银行：钱怎么还；政府：为什么值得支持；董事会：为什么值得做、要什么资源）。

## 文件地图

```
bankable/
├── SKILL.md                         本文件：流程与规则
├── references/
│   ├── intake-questions.md          访谈问题库与追问思路
│   ├── research-and-evidence.md     登记册格式、证据分级、检索与引用、不编造清单
│   ├── audience-profiles.md         四类读者的评审逻辑、章节权重、致命错误、多读者策略
│   ├── chapter-guide.md             总目录、每章必须回答的问题与常见错误
│   ├── financial-model-guide.md     模型结构、假设字段说明、情景设计、指标解读、检查项
│   ├── review-checklist.md          红队评审问题库、一致性与写作质量清单、评审备忘格式
│   ├── pitch-deck-guide.md          路演 PPT 结构（按读者）、每页内容来源、讲稿
│   ├── bilingual-guide.md           英文版改写原则、术语表、数字与政策名称处理
│   └── industry-modules.md          行业画像方法 + 常见行业的驱动因素/指标/风险模块
├── scripts/
│   ├── build_financial_model.py     assumptions.json → 带公式的 Excel 三表模型
│   ├── recalc_and_extract.py        LibreOffice 重算、三情景、导出 model_outputs.json、写回情景对比
│   ├── run_sensitivity.py           单因素敏感性：表格、龙卷风图、JSON
│   ├── make_charts.py               从 model_outputs.json 生成财务图表 PNG
│   ├── build_docx.py                plan.md → 排版好的 Word（中/英）
│   ├── build_pptx.py                deck.md → PPT（无 pptx skill 时的备选）
│   └── check_consistency.py         正文数字 vs 模型输出核对
└── assets/
    ├── assumptions-example.json     假设文件示例（虚构案例，含每个字段的说明）
    ├── plan-skeleton.md             章节骨架与写作提示
    ├── register-template.md         登记册模板
    └── review-memo-template.md      评审备忘模板
```
