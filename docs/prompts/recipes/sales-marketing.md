# Sales and marketing recipes

## 1. Sales dashboard in one call

```text
Use the slim-sheet tools only. <path>, sheet <orders sheet>. overview, then build_report with
title "Sales <period>", KPIs "Revenue=sum:<revenue col>", "Orders=count", "Avg order=avg:<revenue col>|currency",
"Customers=count:<customer col>", and charts "trend:<month col>:<revenue col>",
"rank:<region or rep col>:<revenue col>", "share:<channel col>:<revenue col>".
Before building, query every number you will put in a chart title. Report each KPI value.
```

## 2. Top customers and their share

```text
Use the slim-sheet tools only. In <path>, query revenue per customer for <period>, sorted
descending, with running share of total. Write the top 20 to a sheet "Top customers" (rank,
customer, revenue, share, cumulative share), style_sheet(preset="report"), rules(kind="data_bar")
on revenue. Report how many customers make 80% of revenue.
```

## 3. Commission calculator

```text
Use the slim-sheet tools only. <path> has sales per rep. Commission plan: <e.g. 3% up to 100,000,
5% above; bonus 1,000 when target is met>. Targets per rep: <list or sheet>.
Write a sheet "Plan" with the rates and thresholds as labelled input cells (format_range role="input").
Write a sheet "Commission": rep, sales (SUMIFS on the data), target, attainment %, commission
(formula on Plan cells), bonus, total. Fill down, style_sheet(preset="finance", total_row=true).
evaluate two reps by hand-checkable cases and report them. what_if(op="delta",
inputs=["Plan!<rate cell>=<new rate>"], outputs=["Commission!<total cell>"]) to show the cost of a change.
```

## 4. Pipeline by stage

```text
Use the slim-sheet tools only. <path> is a CRM export (deal, owner, stage, amount, probability,
close date). query count, amount and weighted amount (amount x probability) per stage and per
owner for deals closing in <period>. pivot(rows=["<stage>"], values=["<amount>"], native=true,
to_sheet="Pipeline pivot") for a refreshable view. Flag deals with a close date in the past that are
still open (query) on a sheet "Overdue deals".
```

## 5. Campaign ROI

```text
Use the slim-sheet tools only. Spend file <spend path> (campaign, channel, spend) and results file
<results path> (campaign, leads, customers, revenue). merge on campaign (how="left" from spend);
add CPL = spend/leads, CAC = spend/customers, ROAS = revenue/spend (IFERROR) as formulas; style_sheet;
add_chart(type="scatter") with spend on x and revenue on y, titled with the best and worst ROAS
campaigns (from a query). List campaigns with spend but no results.
```

## 6. Update a price list

```text
Use the slim-sheet tools only. In <path>, sheet <prices>, raise prices of category <X> by <y%>,
rounded to <0 or 2> decimals, keep other categories unchanged. Add a column "New price" with
=ROUND(<price cell>*(1+<y%>),<d>) only on <X> rows (IF on the category), fill down, copy formats
from the price column (range_ops op="copy", formats_only=true). diff the working copy against the
source and confirm only the new column changed.
```
