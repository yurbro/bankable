# Sample: North Yorkshire dairy farm — bank loan with Defra grants (fictional, English / GBP)

Input: a two-year accounts summary CSV and one paragraph — a 440-cow family partnership wanting a 600-cow robotic unit (~£3.1m), applying for a £2.2m 15-year loan with possible Defra capital grants. The user was not available for questions.

| File | Notes |
|---|---|
| `North-Yorkshire-Dairy-Expansion-Business-Plan-2026-11.docx` | 44-page lender version: loan proposal and debt service first, DSCR by year, milk-price sensitivity, with/without-project comparison, gates |
| `North-Yorkshire-Dairy-Expansion-Financial-Model-2026-11.xlsx` | £k three-statement model (English labels), opening balance sheet from the CSV, existing loan + two new tranches + working-capital facility, grants as non-equity capital |
| `…-Assumptions-and-Evidence-Register-2026-11.md` | 22 F / 29 A / 8 E / 10 T items and 22 sources (AHDB, Defra, Bank of England, grant schemes) |
| `…-Review-Memo-2026-11.md` | 20 red-team questions (credit officer, CFO, investor, grant reviewer) |
| `delivery-note.md` | the reply handed to the user: assumptions to confirm, DSCR, minimum cash, limitations |
| `plan.md`, `assumptions.json`, `model_outputs.json`, `sensitivity_outputs.json`, `consistency_report.md`, `charts/` | working files; `assumptions.json` is a useful template for an existing UK business |

Key finding the run surfaced on its own: the brief's £2.2m / 15-year structure fails a prudent 40ppl test (DSCR 1.05, negative cash 2029–32); the plan proposes a 20-year term with a capital holiday, staged drawdown and a working-capital facility, and states that the economics are marginal (break-even 38.6ppl at DSCR 1.0).

The farm, its accounts and all figures are fictional; market data was looked up on 2 Nov 2026 and is unverified.
