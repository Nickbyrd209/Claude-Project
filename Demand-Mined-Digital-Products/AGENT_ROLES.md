# The Six Agent Roles

**How they actually ran:** In this first execution, the six roles ran **one after another inside a single AI session**. I did not launch separate autonomous agents. Each role has its own responsibilities, deliverable folder and quality gate, so the work can later be handed to separate agents or scheduled sessions without being reorganized.

| # | Role | Owns | Deliverables (folder) | Quality bar | Reports to |
|---|---|---|---|---|---|
| 1 | **Market Intelligence** (Head of Market Research) | Finding niches with demand evidence, sourced and dated | `01_Market_Research/` | Every claim labeled VERIFIED / PROXY / ESTIMATE / HYPOTHESIS, with a link | Agent 6 |
| 2 | **Pain Point & Gap Analysis** (Head of Product Opportunity) | Complaints, gaps, 100-point scoring, GO/CONDITIONAL/NO-GO | `02_Competitive_Analysis/`, `03_Opportunity_Scoring/` | Scores explained with evidence. Includes a disproof test for each | Agent 6 |
| 3 | **Digital Product Development** (Head of Product Engineering) | Spec, MVP, build scripts, user instructions | `04_Product_Development/` | Original work. Formulas, not hard-coded values. Reproducible build | Agent 6 |
| 4 | **Marketing & Content** (CMO) | Ideal customer, value proposition, demo concepts, scripts, calendar, message tests | `06_Marketing/` | No fake testimonials, invented outcomes or false urgency. Nothing posted without approval | Agent 6 |
| 5 | **Sales & Conversion** (CRO) | Platform comparison, pricing, product page, upsells, refunds, reporting, unit economics | `07_Sales/`, `08_Financials/` | Revenue kept separate from profit. Documented assumptions | Agent 6 |
| 6 | **Operations & QA** (COO / Quality Controller) | Pipeline, QA, costs, sources, compliance, approvals, reports | `05_Product_QA/`, `09_Management_Dashboard/`, `10_Weekly_Reports/` | Nothing is called "tested" or "done" unless it actually happened. Enforces approval gates | Owner |

**Approval boundary (all roles):** Research, analysis, internal prototypes, drafts and calculations go ahead without sign-off. Spending, accounts, publishing, posting, external messages, credentials, real customer data and legal commitments → **owner approval first** (`09_Management_Dashboard/Approvals_Needed.md`).
