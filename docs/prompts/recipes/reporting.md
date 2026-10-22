# Reporting recipes

## 1. One-page KPI dashboard

```text
Use the slim-sheet tools only. From <path>, sheet <data>, build_report(title="<subject> <period>",
subtitle="<source, date>", kpis=[4-6 items], charts=[4 items with different intents: rank, share,
trend, relationship or compare]). Before the call, query each number you plan to put in a title.
After the call, read the KPI cells with evaluate and report them.
```

## 2. Monthly management pack

```text
Use the slim-sheet tools only. Build <out path> with sheets Summary, Dashboard, Revenue, Costs,
Cash, Data Quality, Notes from <sources>. Summary: 5 findings (claim + number + comparison +
sheet link via annotate op="link" url="#'Sheet'!B4"). Each detail sheet: question in B1, answer in
B2, tables styled with style_sheet, at most one chart (add_chart with an insight title).
Hide helper sheets, Summary first, landscape fit-to-width on every sheet.
```

## 3. Data for board slides

```text
Use the slim-sheet tools only. Prepare small tables (max 8 rows each) that a slide can show:
<list the 3-5 questions>. For each, query the numbers, write them to a sheet "Slide <n>" with a
one-sentence headline in B1, style_sheet(preset="minimal") and add_chart with the headline as
title. Report each headline with its query.
```

## 4. Period-over-period comparison

```text
Use the slim-sheet tools only. In <path>, compare <period A> with <period B> by <dimension>.
query both periods, the change and the change % per <dimension> and in total; sort by absolute
change. Write a sheet "Comparison", style_sheet, rules(kind="color_scale") on change %, and
add_chart(type="bar") of the change. The sum of the changes must equal the total change: show both.
```

## 5. Executive summary in chat only

```text
Use the slim-sheet tools only and change nothing. Read <path> (overview, profile, query) and
answer in at most 8 lines: the 3 most important numbers with their comparison, the biggest risk
in the data (anomalies), and one question to ask next. Show the SQL for each number at the end.
```
