# QA Report: Home Daycare Tax-Ready Records Workbook (prototype v0.1)

**Agent:** Business Operations & QA · **Date:** 2026-11-08

## What was actually tested
| Test | Method | Result |
|---|---|---|
| Formula evaluation | Recalculated in LibreOffice (headless). All 1,350 formulas checked | **0 errors** |
| Calculation correctness | `test_workbook.py` fills copies with synthetic test data across 5 scenarios and compares each output against a separate Python re-implementation of the same rules | **194 / 194 checks passed** (`automated_test_results.md`) |
| Scenario A | Full year, 4 families, 50 working weeks, 168 expenses, 12 months of meals, 80 trips across the July rate change | Pass. Time-Space % = 28.72%, net profit $29,832.26, SE tax $4,215.16 |
| Scenario B | Food method switched to Actual Food Cost | Pass. Food receipts count, meal allowance row goes to $0 |
| Scenario C | Home square footage = 0 (divide-by-zero guard) and a loss year | Pass. Space % = 0, no errors, SE tax = $0 |
| Scenario D | Very high income | Pass. The 12.4% portion is capped at the $184,500 wage base |
| Scenario E | Partial year (100 days) and a misspelled family name | Pass. Time % capped at 100%, "Family not found" warning appears |
| Hand check | SE tax for scenario A: 29,832.26 × 0.9235 × 0.153 = 4,215.16 | Matches |

## Not yet tested (honest gaps)
| Gap | Why | Who / how |
|---|---|---|
| Opening in **Microsoft Excel** desktop and web | Not available in this environment | Owner: open the file, try the dropdowns, print one Parent Statement (5 min) |
| Import into **Google Sheets** | Needs a Google account | Owner: File → Import. Check that dropdowns and the red highlighting still work |
| Visual layout, print preview, mobile view | No visual rendering was checked | Owner, or a later session with screenshots |
| Testing with a **real provider** (nontechnical) | Requires outreach (approval needed) | Validation plan, step 2 |
| **Professional tax review** of the rules and wording | Requires paying a CPA/EA (approval needed) | Budget line B4 |
| Accessibility | Colors were chosen for contrast, but there has been no screen-reader test. Spreadsheets have limited accessibility in general | Add an accessibility note to the product page |

## Known limitations (to disclose on the product page)
- Meal rates are the Jan 1 2026 CACFP Tier I rates for the contiguous states. Alaska and Hawaii users must edit them.
- The 76¢ mid-year mileage rate comes from secondary reporting of IRS Announcement 2026-11 and must be checked on irs.gov before launch.
- Form 8829 limit: flagged only, not calculated.
- The SE tax estimate ignores W-2 wages from another job, and the income tax set-aside is a percentage the user chooses.

## Compliance checklist
- [x] Disclaimer: "record-keeping tool, not tax advice" on Start Here and Tax Summary
- [x] No copyrighted competitor content used. All wording and structure are original
- [x] Example data is fictional and marked EXAMPLE
- [ ] CPA/EA review before sale (approval required)
- [ ] Verify the 76¢ rate and the IRS meal-rate rule on irs.gov before publishing
- [ ] Check whether Etsy's Oct 2026 AI/creativity standards require disclosure that AI was used, before listing on Etsy

**Launch-readiness verdict:** The prototype is internally consistent. It is **not ready to sell** until the Excel and Google Sheets checks, the professional review, and customer validation are done.
