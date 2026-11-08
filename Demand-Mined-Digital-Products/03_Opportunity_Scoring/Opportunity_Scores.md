# Opportunity Scoring: 15 Niches

**Agent:** Customer Pain Point & Gap Analysis (Head of Product Opportunity) · **Date:** 2026-11-08
**Data file:** `opportunity_scores.csv` is the source of truth. The management dashboard reads from it.

**Weights:** Demand 25 · Severity 20 · Gap 20 · Ease 15 · Acquisition 10 · Margin 10.

**Penalty rule:** Where I found no verified sales data, demand is capped at 15/25. Only *budget spreadsheets* (verified shop-level sales counts) and *freelancer tax* (verified population size, unverified purchasing) score above 15. Every score is an analyst judgment based on the evidence in `01_Market_Research/Niche_Research_Report.md`. None of them is a measurement.

| # | Niche | Dem /25 | Sev /20 | Gap /20 | Ease /15 | Acq /10 | Mrg /10 | **Total** | Confidence |
|---|---|---|---|---|---|---|---|---|---|
| 1 | Family child care tax & records | 15 | 16 | 15 | 12 | 6 | 9 | **73** | Medium-Low |
| 2 | Freelancer quarterly tax set-aside | 17 | 15 | 6 | 13 | 7 | 8 | **66** | Medium |
| 3 | Cleaning business pricing & job profit | 13 | 14 | 10 | 12 | 8 | 9 | **66** | Low |
| 4 | Personal budget spreadsheets | 23 | 8 | 3 | 13 | 6 | 8 | **61** | High demand / extreme saturation |
| 5 | Handyman / contractor estimating | 12 | 15 | 9 | 9 | 6 | 9 | **60** | Low |
| 6 | Family caregiver binder | 10 | 15 | 8 | 12 | 7 | 8 | **60** | Low |
| 7 | Landlord / Schedule E tracker | 15 | 12 | 5 | 12 | 6 | 9 | **59** | Medium |
| 8 | Short-term rental host tracker | 13 | 11 | 7 | 12 | 6 | 9 | **58** | Low |
| 9 | Mobile notary / signing agent | 11 | 10 | 7 | 13 | 6 | 9 | **56** | Low |
| 10 | PTA / PTO treasurer kit | 9 | 13 | 9 | 12 | 4 | 9 | **56** | Low |
| 11 | Gig / delivery driver tracker | 12 | 12 | 5 | 12 | 6 | 8 | **55** | Low |
| 12 | Debt payoff planner | 14 | 11 | 3 | 14 | 6 | 7 | **55** | Medium |
| 13 | Wedding budget planner | 13 | 9 | 4 | 13 | 6 | 9 | **54** | Medium |
| 14 | Nonprofit grant & donor tracker | 9 | 12 | 4 | 11 | 4 | 9 | **49** | Low |
| 15 | Job application tracker | 8 | 11 | 2 | 14 | 6 | 7 | **48** | Medium |

Freelancer tax and cleaning are tied at 66. Freelancer tax ranks higher on the tie-breaker of stronger demand evidence. Cleaning has the better gap.

## Why the top scores are what they are

**#1 Family child care — 73**
- *Demand 15/25 (PROXY):* People clearly pay for help here. A paid record-keeping guide is in its 9th edition, a software category sells at about $99/yr, and Etsy listings are starting to appear. Capped at 15 because I could not verify any sales volume.
- *Severity 20/20 → 16:* The time-space percentage changes how much of rent or mortgage interest, utilities, insurance and taxes can be deducted. Picking the standard meal allowance over receipts changes a deduction worth thousands of dollars. Errors cost real money. Not scored 20 because I have no first-hand complaints from providers.
- *Gap 15/20:* Generic "Schedule C ledgers" don't calculate time-space or meal allowances and don't produce per-family statements. Software is a subscription. The book is manual work. The gap is specific and easy to demonstrate.
- *Ease 12/15:* All the logic is public (IRS Pub. 587, Form 8829, CACFP rates). A spreadsheet suits the job. Points off for tax-adjacent liability and the need to update rates every year.
- *Acquisition 6/10:* Buyers gather in identifiable places (Facebook groups, state networks, CACFP sponsors). Those groups need admin permission to post in, and the total market is only about 100k programs.
- *Margin 9/10:* Digital delivery. A $25–35 price is supported by the book's $20–22 price for a non-interactive product and the software's $99/yr.

**#2 Freelancer tax — 66.** The market is huge, but the gap is close to zero: good free calculators and dozens of near-identical sheets already exist. We would be competing on price, which the brief rules out.

**#3 Cleaning pricing — 66.** The pain is real and the "you're actually earning $11/hour" moment is very easy to show on video. Demand evidence is thin (low-priced listings), and free vendor calculators exist.

## What would disprove each top hypothesis
- **Child care:** Fewer than 3 of 10 interviewed providers say they find the time-space or meal-allowance calculation confusing, *or* most of them already pay a preparer who handles everything from shoeboxes, *or* fewer than 1% of qualified landing-page visitors buy at $19 or more.
- **Freelancer tax:** Buyers see no difference from free calculators (expected).
- **Cleaning:** Owners say they set prices by "what the market pays" and wouldn't use a cost calculator.

**Recommendation for further validation:** #1 Family child care (pilot), #3 Cleaning pricing (second candidate), #2 Freelancer tax (watch list only, NO-GO for now).
