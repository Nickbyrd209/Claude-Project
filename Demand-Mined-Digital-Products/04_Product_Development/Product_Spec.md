# Product Specification: Home Daycare Tax-Ready Records Workbook (2026 edition)

**Agent:** Digital Product Development · **Status:** Prototype built and auto-tested. Not yet validated with customers.
**Files:** `Home_Daycare_Tax_Ready_Workbook_2026.xlsx` (the product) · `build_workbook.py` (reproducible build) · `cell_map.json` (cell addresses used by QA)

## 1. Target customer
A U.S. licensed or registered **family child care provider** who runs child care in their own home, has 3 to 12 children, files Schedule C, and keeps records on paper, in a notebook, or in a basic spreadsheet. They aren't a bookkeeper and they're short on time. Many give their records to a paid preparer in January or February.

## 2. Problem solved
At tax time, the provider (or their preparer) needs:
1. A **Time-Space Percentage**: business hours ÷ hours in the year, multiplied by daycare square footage ÷ home square footage. This share applies to rent or mortgage interest, utilities, insurance and property tax.
2. A **food method decision**: the IRS standard meal and snack rates, or actual food receipts.
3. **Business mileage** at the right IRS rate (2026 had a mid-year change).
4. Expenses sorted into Schedule C categories, with shared items reduced by the Time-Space %.
5. A **year-end payment statement for every family**, which parents use to claim the child care credit.
6. A rough idea of **how much to set aside**.

Generic ledgers do #4. Nothing low-cost we found does #1, #2, #5 and #6 together.

## 3. Essential features (MVP) — all built
| Tab | What it does |
|---|---|
| Start Here | Plain-language instructions, a color legend and the disclaimer |
| Setup | Home and daycare square footage, days available, food method, rates with sources, set-aside % |
| Families | One row per family. Paid-by-parent and subsidy totals fill in automatically |
| Time Log | 53 weeks. Hours with children plus prep/cleanup hours. Annual total feeds the Time %. A red flag appears if a week exceeds 168 hours |
| Income | Payments by family and type (Parent Payment / Subsidy / CACFP / Other) using dropdowns |
| Expenses | Category and use dropdowns. "Counts toward deduction" calculates automatically: 100%, × Time-Space %, $0 for personal, and $0 for food receipts when the standard allowance is chosen. Rows with a missing category turn red |
| Meals | Monthly breakfast, lunch, supper and snack counts multiplied by the standard rates |
| Mileage | Trip log. The rate switches automatically on Jul 1 (72.5¢ → 76¢) |
| Tax Summary | Percentages, income, Schedule C-style categories, Form 8829-style home costs, a check against the home-expense limit, estimated net profit, a food method comparison, and a SE tax and quarterly set-aside estimate |
| Parent Statements | Choose a family from the dropdown to get a print-ready statement (provider info, parent, child, total paid, subsidy paid separately, and a signature line). Warns if the family name isn't found |

## 4. Explicitly out of scope (MVP)
- Depreciation calculations (entered manually from the preparer), the exclusive-use rule, and actual-cost vehicle method details
- Carrying the Form 8829 limit forward. The workbook only *flags* when the limit applies
- Attendance, billing and parent messaging (deliberately left out: we are not building software)
- State-specific rules, and Alaska/Hawaii meal rates (users can edit the Setup rates)

## 5. Design principles applied
- **Formulas, not hard-coded values.** 1,350 live formulas. Every rate is a labeled input with its source.
- **Excel and Google Sheets compatibility.** Only functions that work in both (SUMIFS, COUNTIFS, INDEX/MATCH, IF, MIN/MAX, IFERROR). No XLOOKUP, FILTER, macros or scripts.
- **Nontechnical users.** Dropdowns, example rows, yellow input cells, and plain-English labels ("How is it used?" instead of "allocation basis").
- **Honest framing.** This is a record-keeping and organizing tool, not tax advice. Pub. 587 and Form 8829 are cited throughout.

## 6. Yearly maintenance (repeatable)
Every November, update 6 cells on Setup (meal rates, two mileage rates, the wage base and the tax year), rebuild with `build_workbook.py`, rerun the QA script, and publish the "2027 edition". This takes about 1 hour a year.

## 7. Roadmap (only if validated)
1. Free lead magnet: a one-tab "Time-Space % Calculator".
2. A Google Sheets "copy link" version, if interviews show a preference for it.
3. Complementary bundle: parent contract and policy templates, enrollment packet, CACFP meal-count sheets.
4. An adapted version for in-home businesses with similar rules (e.g., home-based salons). Not researched yet.
