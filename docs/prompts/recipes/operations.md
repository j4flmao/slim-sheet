# Operations recipes

## 1. Inventory reorder list

```text
Use the slim-sheet tools only. <path> has item, on hand, on order, average daily usage, lead time
(days) and safety stock. Add Reorder point = usage x lead time + safety stock and Order qty =
MAX(0, reorder point + usage x <review days> - on hand - on order) as formulas (write + fill).
range_ops(op="copy", where="\"<order qty col>\" > 0", to="Reorder!A1", values_only=true) after Excel
values exist, or query the same condition and write the result. style_sheet; report how many items
need ordering and the total quantity.
```

## 2. Supplier on-time-in-full (OTIF)

```text
Use the slim-sheet tools only. <path> has purchase order lines (supplier, promised date,
delivered date, ordered qty, delivered qty). query per supplier: lines, on time %, in full %, OTIF %
for <period>. Write "Supplier OTIF", style_sheet, rules(kind="color_scale") on OTIF %,
add_chart(type="bar", title=<best and worst supplier>). Report suppliers below <90%>.
```

## 3. Stock aging

```text
Use the slim-sheet tools only. <path> has item, receipt date, quantity, unit cost. query value
by age bucket (0-30, 31-90, 91-180, >180 days at <date>). Write "Stock aging" and a share chart
(build_report charts=["share:<bucket>:<value>"] on that sheet or add_chart type="doughnut").
Report the value older than 180 days and its share.
```

## 4. Order backlog

```text
Use the slim-sheet tools only. <path> has open orders (order date, promised date, customer,
value, status). query backlog value and count by promised week, and orders already late at
<date>. pivot(rows=["<customer>"], values=["<value>"], to_sheet="Backlog by customer").
Report the late value and the 5 largest late orders by value.
```

## 5. Production plan vs actual

```text
Use the slim-sheet tools only. Plan <plan path> and actual <actual path> by line and day.
merge(how="full", keys=["<line>","<date>"]); add Attainment = actual/plan (IFERROR), fill down;
style_sheet; rules(kind="less_than", value="0.9") on attainment. add_chart(type="line") of plan
and actual per day for the total. Report lines under 90% and days with production but no plan.
```
