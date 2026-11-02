# North Yorkshire Dairy Expansion — Review Memo

Review date: 2026-11-02　Perspectives: agricultural lender / credit committee (primary); partners as equity holders; grant assessor; CFO-style challenge　Subject: Business Plan V1.0 draft (plan.md → .docx), Financial Model 2026-11 (.xlsx, English and original), register V1.0

## 1. Consistency check result

`check_consistency.py plan.md model_outputs.json --extra sensitivity_outputs.json incremental_outputs.json plan_facts.json model_outputs.json`: all eight required figures present (2026 and 2035 revenue, 2035 net income, external funding, peak funding need, 2035 gross margin, minimum operating DSCR 1.17, downside result). 276 unit-bearing numbers scanned in the financial chapters, 48 skipped as register/source-tagged, **one flagged**: `107.5k` in section 11.8 — this is the downside 2027 closing cash of −£107.5k, present in `model_outputs.json` as a negative value; the heuristic scan does not match signs. Verified manually; no change needed.

Model checks (`recalc_and_extract.py`): balance sheet balances in all years and scenarios (max difference 0); base case has no negative cash and no year below the £50k floor; warnings retained and addressed in the text: DSCR below 1.2 in 2029 (1.19) and 2031 (1.17) — disclosed in 12.3 with a covenant proposal; downside funding gap £657.0k — chapter 13 and gates; "sources vs feasibility-basis total investment" and "use of funds ≠ sources" warnings — these arise because the skill's T7 table is built for greenfield projects (it counts the whole opening working capital as investment) and because £200k of the uses are met from internal cash; explained in 12.2 and the T7 table was not used.

## 2. Issues raised (Q-xx; risk matrix uses R-xx; register uses F/A/E/T)

| ID | Perspective | Question | Severity (fatal / material / minor) | Resolution | Chapter / model field | Status |
|---|---|---|---|---|---|---|
| Q-01 | Credit | Base-case DSCR is 1.19 in 2029 and 1.17 in 2031 — below the 1.25 most lenders want, and 2028 is interest-only. Is the plan bankable? | Material | Structure changed from 15 to 20 years with 24-month holiday (A-18); working capital facility added (A-21); 25-year term shown to lift the minimum to 1.22; covenant proposal (annual test, cure period) in 12.3; break-even 38.6/40.3ppl stated in 11.7 | Ch.12.3, 12.6; `financing.loans` | Resolved in structure; thin cover disclosed to user |
| Q-02 | Credit | The downside (a 2026 repeat during construction) leaves a £657k gap and even the mitigation variant leaves £415k. What happens then? | Fatal if unaddressed | Gates G2/G3 make construction conditional on price, grants and funds; mitigation package quantified (Table 22); deferral is the primary response pre-construction; partners' support / asset sale disclosed as last resort | Ch.10.2, 11.8, 13.1 | Disclosed; residual risk accepted by design |
| Q-03 | Credit | Where do the own funds come from? Business cash is £190k and 2026 EBITDA only c.£155k. | Fatal if not evidenced | £350k partners' capital assumed (A-20) and listed first among confirmations; proof of funds is gate G3; T-09 | Ch.12.2; `financing.equity` | Awaiting user |
| Q-04 | Credit | £3.1m for 600 places with 10 robots and slurry infrastructure is £5,167 per place against £7,500–10,000 benchmarks. Credible? | Material | Flagged in E-04 and the register's disagreement log; capex +15% sensitivity (£479k gap); G1 requires fixed-price quotes plus 5–10% contingency; T-03 | Ch.8.5, 13 (R-03) | Disclosed; awaiting quotes |
| Q-05 | Credit / grant assessor | Grants of £300k are assumed although SIG, IFP and FETF rounds are closed at Nov 2026. | Material | Scheme status set out with sources (4.4); "no grant" case (£284k gap) and alternative funding plan (12.5); grant offer is a G2 condition; T-07 | Ch.4.4, 12.5 | Disclosed |
| Q-06 | Credit | Partners' drawings assumed inside the £51.5k family labour charge; tax at 28% with no capital allowances. Both could be wrong in opposite directions. | Material | Drawings +£50k sensitivity (DSCR 1.02); AIA noted as unmodelled upside; T-05 asks for partner numbers and drawings | Ch.11.1, 11.8, 13 (R-11) | Disclosed |
| Q-07 | Credit | Security and loan-to-value: 870 acres owned but no valuation; existing lender's charge. | Material | `[to be provided — T-04]` placeholders; E-08 debt per cow; no value invented | Ch.12.7 | Awaiting user |
| Q-08 | Credit | 2026 trading is an estimate; the bank will want management accounts. | Material | A-01/A-02 flagged; T-01; 2026 column labelled estimate | Ch.11, 11.9 | Awaiting user |
| Q-09 | Credit | Milk contract: can the processor take 5.8m litres, and on what price mechanism? | Material | T-06; G2 requires written confirmation | Ch.5, 7, 10.2 | Awaiting user |
| Q-10 | Credit | Interest in 2027 assumes tranche 1 fully drawn from March and the facility fully drawn in 2028 — conservative; are other assumptions equally conservative? | Minor | Noted in 12.4; constant-price basis and nominal interest also conservative; yield +7.9% vs 10% guideline | Ch.11, 12.4 | Disclosed |
| Q-11 | CFO | The incremental IRR is c.0–1.5% and incremental NPV negative; why invest? | Material | Stated plainly in 1 and 11.6; the status-quo comparison ignores slurry compliance and parlour renewal (A-27); strategic drivers (labour, compliance, asset life) set out; scaled variant run (no DSCR benefit) | Ch.1, 3.4, 11.6 | Disclosed; user decision |
| Q-12 | CFO | Constant 2026 prices for both revenue and costs — what if milk price lags cost inflation as in 2026? | Material | ±3ppl and ±10% variable cost sensitivities; break-even analysis; downside scenario | Ch.11.7, 11.8 | Disclosed |
| Q-13 | CFO | Robots have a 12-year life and will need replacing c.2040, inside the 20-year loan term. | Minor | Sentence added to 12.6 recommending a replacement reserve from post-2032 surpluses or asset finance at that point | Ch.12.6 | Resolved |
| Q-14 | Grant assessor | Current slurry storage months, EA notification and planning status unknown; 6-month storage is the scheme standard. | Minor | T-03, T-08 placeholders; no compliance status invented | Ch.2.3, 3.3, 8.6 | Awaiting user |
| Q-15 | Partners | A single-price business with peak debt of c.£4,800 per cow; land sale is the ultimate backstop. | Material | Stated in 13.1; gates to avoid starting into a trough | Ch.13 | Disclosed |
| Q-16 | Writing quality | Chapter 9 (management) is almost empty. | Minor | By design — only user-supplied facts may appear; placeholders and T-05 | Ch.9 | Accepted |
| Q-17 | Model | `use_of_funds` total (£3,350k) ≠ external sources (£3,150k). | Minor | Explained: £200k from internal cash (E-07); Table 18 balances | Ch.12.2 | Resolved |
| Q-18 | Model | Grants modelled as a capital reserve rather than deferred income. | Minor | Footnote in 11.5; no cash or DSCR effect | Ch.11.5 | Resolved |
| Q-19 | Model | Capital allowances (AIA on c.£1.3m of plant) not modelled. | Minor | Tax is already nil 2028–35 under the model's loss carry-forward, so cash effect is limited to 2026–27 (c.£7k); noted as upside | Ch.11.3 | Disclosed |
| Q-20 | Model | Skill's Excel is Chinese-labelled. | Material for a UK user | English version produced by `translate_model_en.py` and verified cell-by-cell against `model_outputs.json`; the original is kept for re-running the skill scripts | Deliverables | Resolved |

## 3. Matters still to be disclosed to the user after revision

1. The recommended structure differs from the brief: 20-year term (not 15) and a £350k working capital facility; the 15-year version fails the plan's own tests at 40ppl.
2. The project is marginal at a prudent price: DSCR 1.17–1.34 in the first decade; break-even 38.6ppl; incremental return c.0–1.5% on a ten-year view.
3. Grants are uncertain; the plan should not go to the bank without a grant offer or the agreed alternative.
4. Own funds of c.£350k must be evidenced.
5. 2026 trading figures, cost splits, quotes, contract and land valuation are needed before submission (T-01 to T-10).

## 4. Assumptions the user should confirm first (largest effect)

| Register ID | Assumption | Current value | Effect of a ±20% (or stated) change on the key conclusions |
|---|---|---|---|
| A-03 | Milk price from 2028 | 40.0ppl | ±3ppl: cumulative net income +£1,098k / −£1,465k; −3ppl: £1,114k gap, DSCR 0.65 |
| A-07 | Variable cost per cow | £2,865 | ±10%: NPV ±£959k; +10%: £1,450k gap |
| A-20 | Partners' capital introduced | £350k | Any shortfall is a direct funding gap in 2027–29 |
| A-10 | Capital cost | £3,100k | +15%: £479k gap |
| A-19 | Grants | £300k in 2028 | None: £284k gap |
| A-18 | Loan term / holiday | 20 years / 24 months | 15 years: DSCR 1.05; no holiday: DSCR 0.85 |
| A-08 | Private drawings | within £51.5k family labour charge | +£50k a year: DSCR 1.02, £230k gap |
| A-05 | Yield on robots | 9,600 l/cow | ±5%: DSCR 1.11 / 1.24 |
| A-04 | Herd build-up timing | 600 cows by 2030 | 12-month delay: DSCR 0.99, £60k gap |
| A-01/A-02 | Year-end; 2026 estimate | Calendar year; 37.0ppl | Changes opening cash and own-funds availability |
