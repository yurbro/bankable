# Rotterdam Textile-to-rPET Plant — Review Memo (red team)

Review date: 2026-11-02　Perspectives: equity investor (primary), grant evaluator (secondary), credit officer and CFO/board (as the plan carries €6.6m of debt and a gated Series B)　Object: `plan.md` V1.0, `model_outputs.json` (generated 2026-11-02), `sensitivity_outputs.json`

## 1. Consistency check result

`check_consistency.py plan.md model_outputs.json --extra sensitivity_outputs.json` — **all 9 required figures present** (first-year and final-year revenue, final-year net income, loan total, minimum cash, final gross margin, project IRR, operating DSCR, downside revenue). 382 unit-bearing numbers scanned in the financial chapters; 61 skipped as tagged to the register; 15 flagged for manual review, all resolved as follows:

| Flagged | Resolution |
|---|---|
| `1%`, `€60m`, `€634m` | Sourced regulatory/market facts [S-03, S-09, S-10], not forecasts — acceptable |
| `80%`, `92%`, `74%` | Gate condition, utilisation input [A-01], and a derived break-even share (8,900 t ÷ 12,000 t) — acceptable |
| `€90k`, `€60k`, `€95k`, `€80k` | Model inputs listed in the key-assumptions table (opex lines, cost per FTE, opening payables); the model exports category totals only — acceptable |
| `−€347k` | `monthly_year1.difference_vs_annual` in the model output, a field the checker does not index — acceptable |
| `−€486k` | No-grant case minimum cash (−485.7 in `sensitivity_outputs.json`); the checker does not read the Unicode minus sign before a currency symbol, so it compared +486 — acceptable (tool limitation, reported) |

Two narrative figures were corrected during review (working-capital absorption 2027/cumulative; allocation of Series A proceeds) so that every sentence reconciles to the series.

## 2. Issues

| ID | Perspective | Question / issue | Severity | Resolution | Chapter / model field | Status |
|---|---|---|---|---|---|---|
| Q-01 | Investor | "€4m to build a 12 kt plant" does not add up: feasibility-basis investment is €12.9m. An IC will reject a plan that hides this. | Fatal if unaddressed | Plan restructured as a phased build: Line 1 (2027) funded by Series A + €3.6m equipment loan + uncertain grant; Line 2 (2029) by €3.0m loan + €2.5m gated Series B. Stated in §1.1, §12.1–12.3 and the register (E-04, divergence table). | Ch. 1, 12; `financing` | Handled |
| Q-02 | Investor | Unit costs (feedstock €250/t, yield 76%, conversion €300/t) are grade-D estimates; a 10% error halves the IRR. | Fatal if unverified before launch | Cannot be fixed without pilot data. Marked as the first confirmation item (T-02); +10% and +25% unit-cost sensitivities reported; M3 feedstock contracts made a precondition of the Line 1 order. | Ch. 6.3, 11.8, 13 (R-04); A-05 to A-08 | Disclosed |
| Q-03 | Investor | Pellet price €1,250/t has no direct market quote (textile-derived, non-food). −10% price removes the NPV. | Important | Price positioned between published flake and food-grade pellet prices [F-17], with an explicit inference (E-03); virgin-indexed floors made milestone M2; Series A should not close before at least one binding contract. | Ch. 4.3, 6.2, 7.4; A-03 | Disclosed + mitigation |
| Q-04 | Investor | Team chapter is empty. "Who has built and run a polymer line?" | Important | Nothing can be written without the founders' input; chapter states what investors will expect and the three hires required. | Ch. 9; T-05 | For founders |
| Q-05 | Investor | Why will brands pay, given Renewcell? LOIs are non-binding. | Important | §5.3 addresses the precedent directly; LOIs treated as demand evidence only; capacity gated at G2; six-to-eight-customer target. | Ch. 4.6, 5.3, 10.2 | Handled |
| Q-06 | Investor | Equity IRR 37.2% and 7× terminal multiple look generous for a thin-margin recycler. | Moderate | Terminal multiple flagged "to verify" (A-32); IRR without terminal value is undefined within 5 years (cumulative UFCF −€7.8m at 2031) and this is stated; payback beyond horizon stated. | Ch. 11.6, 12.1, 12.5 | Disclosed |
| Q-07 | Credit | Operating DSCR 1.24 in 2029 is at the covenant edge; a 2-year grace on L1 is generous for an equipment loan. | Important | Stated plainly in §11.6 and §12.2 with the expected covenant (cash sweep / equity cure at 1.2×). Alternative (8-year tenor) not modelled — lender feedback (T-09) needed first. | Ch. 11.6, 12.2; A-27 | Disclosed |
| Q-08 | Credit | Debt is sized on un-quoted capex; +20% capex opens a €1.8m gap. | Important | Sensitivity reported; recommendation that lender LTV flex with final quotes; fixed-price equipment contracts in R-05. | Ch. 11.8, 13 | Disclosed + mitigation |
| Q-09 | Credit | 2027 monthly trough (Dec €120k) is far tighter than the annual figure (€467k). | Important | Explained in §11.9 (method difference −€347k); €0.5–0.75m standby working-capital facility recommended alongside the Series A, also covering the no-grant case (€786k gap in 2028). | Ch. 11.9, 12.4, 13.1 | Disclosed + mitigation |
| Q-10 | Grant evaluator | Eligibility is asserted nowhere, but will the project pass DEI+ "new to the Netherlands" and "no start before application"? | Important | §12.4 lists both conditions; M5 sequences the equipment order after filing; the founders must check existing Dutch textile-to-polyester installations against the "new to NL" test before choosing DEI+ as the lead scheme. | Ch. 12.4; F-20 | For founders |
| Q-11 | Grant evaluator | CO₂ and waste-diversion figures are rules of thumb. | Important for any application | Tagged A-40 with "to verify by screening LCA"; chapter 14 says so in its first paragraph. | Ch. 14 | Disclosed |
| Q-12 | Grant evaluator | The 2026 DEI+ round is closed and oversubscribed; EIC is highly competitive — is €1.25m realistic in 2027–28? | Important | Grant modelled as uncertain; no-grant case shows the liquidity effect (€786k gap) and no effect on IRR; recommendation to pursue one lead scheme and arrange standby liquidity. | Ch. 12.4, 11.8; A-26 | Disclosed |
| Q-13 | CFO / board | Opening balance sheet, tax-loss treatment and grant accounting are simplified. | Moderate | A-33 and A-29 flagged; grant booked to capital reserve rather than deferred income (stated §11.5); 2026 accounts to replace A-33 (T-01). | Ch. 11.1, 11.5, 11.9 | Disclosed |
| Q-14 | CFO / board | Downside case is not survivable from operations (€3.7m gap) and deferring Line 2 does not help. | Important | Stated without softening in §11.8; five-step response ladder in §13.1; stop-loss conditions in §13.2. Investors should price this in. | Ch. 11.8, 13 | Disclosed |
| Q-15 | CFO / board | Feedstock concentration: the plant needs more than half of the Dutch polyester-rich stream. | Important | E-05 quantified; multi-country sourcing (M3); R-04 trigger thresholds. | Ch. 4.5, 8.4 | Handled |
| Q-16 | Writing | Chapters 2 and 9 contain many `[to be confirmed]` placeholders. | Moderate | Deliberate (no fabrication); 13 placeholders in the body, listed in the delivery note. | Ch. 2, 9, appendices | Disclosed |
| Q-17 | Writing | Full-width parentheses and Chinese headings leak from scripts into English outputs. | Minor | Sensitivity table taken from `make_tables.py` (ASCII brackets) instead of `sensitivity.md`; consistency report carries an English summary; Excel labels translated in a post-processing step. | Deliverables | Handled |

## 3. Items that remain disclosed to the founders (cannot be resolved in this draft)

1. Unit economics (price, feedstock, yield, conversion) are unverified estimates — the plan's conclusions move materially with them (Q-02, Q-03).
2. The team chapter and the 2026 accounts are missing (Q-04, Q-13).
3. Grant eligibility and timing are uncertain; the plan is shown with and without the grant (Q-10, Q-12).
4. The downside case needs external liquidity; this is a feature of the business's cost structure, not a modelling artefact (Q-14).
5. Terminal multiple, lender terms and equipment quotes are assumptions pending market input (Q-06, Q-07, Q-08).

## 4. Assumptions the founders should confirm first (largest effect on conclusions)

| Register ID | Assumption | Current value | Effect of an adverse move (from sensitivity run) |
|---|---|---|---|
| A-03, A-04 | Selling prices (pellets / flakes) | €1,250/t / €950/t | −10% price: NPV −€3.5m (from +€3.1m), IRR 0.0%, funding gap €2.0m |
| A-01 | Volumes / ramp-up | 2,400 → 11,000 t | −20% volume: NPV −€3.7m, IRR −1.2%, gap €1.9m |
| A-05 to A-08 | Feedstock price, yield, conversion cost | €250/t, 76%, €300/€210 per t | +10% unit cost: IRR 13.6%, gap €0.4m; +25%: gap €3.2m, IRR −6.1% |
| A-19 to A-21 | Capex Line 1 / site / Line 2 | €5.5m / €1.4m / €5.2m | +20% capex: gap €1.8m, IRR 17.7% |
| A-26 | Grant | €1.25m in 2027–28 | No grant: gap €0.8m in 2028; IRR unchanged |
| A-30 | Series B at G2 | €2.5m in 2029 | Not raised but Line 2 built: gap €1.8m; not built: IRR 12.0%, NPV −€0.4m |
| A-27, A-28 | Debt terms | €3.6m + €3.0m, 6.5%, 7 yrs, 2/1-yr grace | Shorter grace or tenor would push 2029 DSCR below 1.2 (not modelled; lender input needed) |
| A-33 | Opening balance sheet | Cash €450k etc. | Lower opening cash reduces the 2027 minimum one-for-one (€467k base) |
| A-17 | Headcount and cost per FTE | 15 → 35 FTE | +10% opex: IRR 18.9%, minimum cash €318k |
| A-32 | Terminal multiple | 7× 2031 EBITDA | Drives the 24.1% IRR; no IRR exists without terminal value within 5 years |
