# Monitoring & Continuous Improvement Workflow (Phase 7)

**Status:** Designed, **not scheduled**. This cloud environment does have a scheduling tool, but nothing has been set up. Scheduling a recurring session needs your approval (Approvals item 12).

## Weekly market review (≈30 min, AI-run; owner reads a 1-page summary)
1. **Competitor watch:** Search Etsy, Payhip and Gumroad for "home daycare tax", "family child care spreadsheet" and "time space percentage". Log new listings, prices and visible review counts in `competitor_log.csv` (create on first run).
2. **Pricing changes:** Check whether KidKare, Brightwheel or Redleaf prices changed.
3. **Rule changes:** Check the IRS newsroom (mileage, Pub. 587 drafts) and USDA CACFP notices.
4. **Customer signals:** Read support emails, Payhip feedback answers and comments on posts (owner provides an export). Tag each one as bug / confusion / feature request / praise.
5. **Sales performance:** Enter weekly actuals on the dashboard *Financial Performance* tab.
6. **New problems:** Run 3 searches in one adjacent segment (rotating list: home-based salons, pet sitters, cleaning businesses, tutors). Add any finding to the opportunity CSV as "Researching".
7. **Output:** `10_Weekly_Reports/Week_YYYY-MM-DD.md` covering what changed, what we learned, the top 3 actions, and approvals needed.

## Monthly strategic assessment (≈1 h)
- Rescore the top 5 niches using new evidence.
- Product decision: **iterate** (fix the top 3 confusions), **expand** (complementary product), or **pause**.
- Review costs and the remaining budget.
- Confirm the WIP limit (1 product in development).

## Ready-to-enable schedule (if approved)
- Weekly: Mondays about 10:40 local time. Run the review above and write the weekly report.
- Monthly: first Monday. Strategic assessment.
- Yearly: November. Run the edition update (`04_Product_Development/build_workbook.py` with new rates), then the QA script.

## Improvement rule
Sales data and feedback go into **improving the existing product first**. A new product is only started when the current one is Active and stable, and the new opportunity scores ≥65 with evidence.
