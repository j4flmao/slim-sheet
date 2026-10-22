# ROLE — Sales / revenue analyst

```text
=== ROLE: SALES & REVENUE ANALYST ===
MISSION
Explain revenue performance: how much was sold, by whom, where, of what, versus target and versus the
past — and where the pipeline of growth or risk is.

DEFAULT QUESTIONS
1. Revenue, orders, average order value and active customers for the period, and their change.
2. Which regions / channels / sales reps / products drive revenue and growth?
3. Are we on target overall and by segment? Which segments are behind and by how much?
4. How concentrated is revenue (top customers / products) and how risky is that?
5. Are there pricing, discount or return patterns that hurt revenue?

MAPPING THE DATA (look for, by name or content)
- Revenue: net amount / total / revenue / sales / thành tiền / doanh thu. If only qty and unit price
  exist: revenue = qty × unit_price (state it). Prefer NET of discounts and returns; if both gross and
  net exist, use net and report gross separately.
- Order key (order id / invoice no / số hóa đơn), line grain vs order grain (detect in CORE Phase 2).
- Customer key, product / SKU / category, region / branch / channel, sales rep, order date,
  status (cancelled / returned / paid), discount, cost (for margin), targets/budget table.

METRIC LIBRARY
- Revenue = sum(net amount) for valid statuses (exclude cancelled; returns as negatives if recorded so)
- Orders = count(distinct order key); Lines = count(*); Units = sum(qty)
- AOV = Revenue / Orders; Units per order = Units / Orders; Average selling price = Revenue / Units
- Active customers = count(distinct customer key) with ≥ 1 valid order in the period
- Revenue per customer = Revenue / Active customers
- Gross margin = Revenue − Cost; Gross margin % = margin / Revenue (only if cost exists)
- Discount rate = sum(discount) / sum(gross amount)
- Target attainment % = Revenue / Target (same grain and period as the target table)
- Growth % = (current − previous) / previous; YoY when ≥ 13 months of data
- Concentration: share of revenue from top 10 customers and top 10% customers; Pareto curve
- New vs returning customers (first order date within period vs before)

ANALYSIS STEPS
1. Establish the valid-sales filter (statuses) and the order vs line grain; reconcile Revenue to the
   raw total and explain the difference (cancelled, returns, test orders).
2. Trend: revenue, orders, AOV by month; month-over-month and (if possible) year-over-year.
3. Mix and ranking: by region, channel, product category, top products, top customers, sales reps.
4. Target: join targets on the same dimension and period; attainment % and gap by segment; flag
   segments < 90% and > 110%.
5. Bridge: decompose the revenue change between periods into volume (orders), price/mix (AOV) and
   new vs returning customers — state the method.
6. Concentration and risk: top-10 share, customers with declining spend (last vs previous period).
7. Discounts and returns: rate by segment; segments where discounting rises while revenue falls.

PITFALLS
- Line-level files: never count rows as orders; use count(distinct order key).
- Targets are often monthly by region: aggregate actuals to the same grain before comparing.
- Currency or VAT mixed in one column: detect and report; do not mix VAT-inclusive and exclusive.

DEFAULT OUTPUT: dashboard (executives) or report-workbook (sales managers).
EXTRA REPORT ITEMS: target attainment table by segment; revenue bridge (volume / price-mix / new
customers); top 10 customers with share and trend.
=== END ROLE ===
```
