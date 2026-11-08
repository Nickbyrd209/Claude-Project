# Unit Economics & Illustrative Scenarios: P-001

**Live model:** `09_Management_Dashboard/Master_Business_Dashboard.xlsx` → *Scenarios (Illustrative)* tab, which calculates with formulas and has editable blue inputs. The figures below were produced by that model after a LibreOffice recalculation and checked by a separate Python calculation.

**These are illustrations, not forecasts.** No sales have happened. Unit volumes are assumptions chosen to show the range.

## Assumptions (documented)
| Input | Value | Source |
|---|---|---|
| Price | $29 | 07_Sales recommendation |
| Payhip fee | 5% | Payhip help center (Free plan) |
| Processor | 2.9% + $0.30 | Approximate Stripe card rate. Verify |
| Software | $0/month | Payhip Free |
| Refunds | 5% / 4% / 3% of gross | Assumption. Fees treated as not returned on refunds |

## Per unit
Gross $29.00 − platform & processing $2.59 = **$26.41 before refunds and ads**.

## Monthly scenarios
| Line | Low | Moderate | High |
|---|---|---|---|
| Units (assumed) | 5 | 25 | 75 |
| **Gross sales** | $145.00 | $725.00 | $2,175.00 |
| Platform & transaction fees | ($12.96) | ($64.78) | ($194.33) |
| Refunds | ($7.25) | ($29.00) | ($65.25) |
| Advertising | $0 | ($50.00) | ($150.00) |
| Software | $0 | $0 | $0 |
| Other direct costs | $0 | $0 | $0 |
| **Contribution margin** | **$124.80** | **$581.23** | **$1,765.43** |
| Contribution margin % | 86.1% | 80.2% | 81.2% |

**Net operating profit before tax** = contribution margin − one-time and overhead costs. If the owner approves the recommended startup budget of $216, a low-volume month 1 would show a net operating **loss** of about $91 ($124.80 − $216). That cost is recovered after about 9 units at $26.41 contribution each, assuming no ad spend.

## Break-even for paid acquisition
With $26.41 contribution per unit, paid ads only make sense if **CAC stays below about $15** (leaving roughly $11 per unit to cover refunds, support and overhead). At an assumed 2% landing-page conversion rate, that means paying no more than about $0.30 per qualified click. **This has not been tested.** The plan is organic-first.

## What these numbers do not include
Income tax, the owner's time (about 3–4 h/week), U.S. sales tax obligations (see the compliance flag in 07_Sales), chargebacks, and currency conversion.
