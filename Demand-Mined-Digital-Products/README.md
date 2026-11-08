# Demand-Mined Digital Products

An AI-run digital product business built on **Demand Mining and Gap-Closing**: find proven demand, identify customer frustration, build a better solution, validate willingness to pay, launch carefully, measure results, and repeat what works.

**Start with:** [`09_Management_Dashboard/Executive_Summary.md`](09_Management_Dashboard/Executive_Summary.md), then [`Approvals_Needed.md`](09_Management_Dashboard/Approvals_Needed.md).

| Folder | Contents |
|---|---|
| `AGENT_ROLES.md` | The six roles, their responsibilities and the approval boundary |
| `01_Market_Research/` | 15-niche research report (evidence labeled) · source register |
| `02_Competitive_Analysis/` | Top-3 pain point and gap analysis with GO / CONDITIONAL / NO-GO |
| `03_Opportunity_Scoring/` | 100-point scores (CSV is the source of truth) |
| `04_Product_Development/` | **Pilot prototype `.xlsx`** · build script · product spec |
| `05_Product_QA/` | QA report · automated test script and results |
| `06_Marketing/` | Validation plan · marketing plan · 30-day `content_calendar.csv` |
| `07_Sales/` | Platform comparison · pricing · draft product page · support/refund workflow |
| `08_Financials/` | Unit economics and illustrative scenarios |
| `09_Management_Dashboard/` | `Master_Business_Dashboard.xlsx` · summary · pipeline · decision log · budget · 30-day plan · approvals |
| `10_Weekly_Reports/` | Weekly reports · monitoring workflow |

## Rebuild & retest
```bash
cd 04_Product_Development && python3 build_workbook.py Home_Daycare_Tax_Ready_Workbook_2026.xlsx
# recalculate with LibreOffice, then:
python3 ../05_Product_QA/test_workbook.py Home_Daycare_Tax_Ready_Workbook_2026.xlsx <path/to/recalc.py> <scratch_dir>
cd ../09_Management_Dashboard && python3 build_dashboard.py
```
Requires Python 3 with `openpyxl`, plus LibreOffice for recalculation. `recalc.py` runs LibreOffice headless and writes the computed values back into the file.

**Status:** No money spent, nothing published, no external accounts or messages. All of those wait for owner approval.
