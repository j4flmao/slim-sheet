# OUTPUT — Dashboard pack (several dashboards, many chart types)

Use when one dashboard is not enough: an executive overview, a "what changed" view, a segment
close-up and a chart gallery, each answering its own question with the chart type that fits it.
Builds on [dashboard.md](dashboard.md); the chart rules are in [design.md §6](../design.md#6-chart-variety-pick-the-chart-from-the-question).

```text
=== OUTPUT: DASHBOARD PACK ===
DELIVERABLE
One .xlsx at OUTPUT_PATH (default "<first source stem>_dashboards.xlsx"), sheets in this order:
  Summary | Dashboard | Growth Dashboard | <Segment> Dashboard | Charts | analysis sheets |
  Data | <Segment> Data | Trend data | helper sheets | Data Quality | Notes
The "<...> Dashboard data" sheets build_report creates are hidden at the end (sheet_ops op="hide").

STEP 1 — DATA SHAPES (charts read stored values, so give each dashboard a value table)
- Data: the cleaned table (convert or write), made an Excel table (set_table) only AFTER the last
  build_report on it, or build_report first and set_table last.
- Trend data (for "what changed"): a LONG table, one row per period x segment:
    Period | Segment | <Measure 1> | <Measure 2> | <Change> (change filled on the latest rows only)
  Build it from query results; at least 2 periods; label periods so they sort (FY2024, FY2025).
- Segment data: range_ops(op="copy", range=<Data block>, to="<Segment> Data!A1",
  where="<SQL condition>", values_only=true).
- Formula columns (SUMIFS, lookups) are fine on analysis sheets, but build_report and add_chart's
  automatic type choice need stored values: point charts at value columns or value blocks.

STEP 2 — DASHBOARDS (build_report, one call each, max 6 KPIs and 4 charts)
Give every dashboard charts with DIFFERENT intents; never two charts with the same message.
  Dashboard (overview)    : rank (biggest items), share (mix of the whole, <= 5 parts -> doughnut),
                            relationship (two measures -> scatter), compare (few groups -> columns)
  Growth Dashboard        : trend:Period:<measure> (x2 measures), rank:<segment>:<change column>
  <Segment> Dashboard     : share, rank (top items), relationship, compare - on the segment data
KPI tiles: totals that must stay live use formula items ('Label==<formula>|format'); ratios use
'Label=ratio:A/B'; never more than 6.
If a chart turns out weak (it sums periods, repeats another chart, has no message): undo(path)
and rebuild without it rather than leaving it.

STEP 3 — CHART GALLERY (add_chart, for types build_report does not pick)
On a "Charts" sheet, one small value block per chart (header row, category column, series
columns), titled with the insight:
  column  : 2+ series side by side (e.g. prior vs latest per segment)
  bar     : long category names or > 8 items, ranked
  line    : one or more series across ordered categories or periods
  area    : cumulative curves (concentration: cumulative share by rank)
  pie     : one whole in <= 6 slices, the main slice highlighted (highlight="<item>")
  doughnut: 2-5 parts
  scatter : first column = x, next = y; say the relationship in the title (with r if computed)
  sparkline: on a table, one per row of numbers (data_range=<row block>, anchor=<first cell>)
Each add_chart gets title="<insight>", x_title / y_title when the units are not obvious.

STEP 4 — FINISH
- rules: data_bar on ranked measures, color_scale on growth, less_than 0 -> bad on changes.
- annotate: a note on every header whose definition is not obvious; in-workbook links from the
  Summary findings to their sheets (url="#'Sheet name'!B4").
- name_ops: names for totals the KPIs and notes refer to.
- sheet_layout: gridlines off and landscape fit-to-width on dashboards and Charts; freeze headers.
- Optional: pivot(native=true) for a refreshable PivotTable of the main measure by one field.

TITLE CHECK (mandatory - a wrong title is worse than no title)
Every number, ranking or comparison in a chart title or KPI label is reproduced by a query BEFORE
the chart is built ("X earns more than Y and Z combined" -> query the three sums; "about 8x" ->
query the ratio; "barely related" -> query corr()). When the check fails, rewrite the title, or
replace the chart with add_chart(replace=true) / rebuild the dashboard.

ACCEPTANCE CRITERIA
[ ] Each dashboard answers a different question; chart intents differ within a dashboard.
[ ] At least 5 different chart types across the pack (bar, column, line or area, pie or doughnut,
    scatter), plus sparklines on the main table when there are two or more periods.
[ ] Every title claim has its verifying query in Notes ("Chart titles" section).
[ ] Tiles that count each entity once vs charts that count rows are labelled (subtitle or Notes).
[ ] Helper data sheets hidden; Summary first; the file opens in Excel without repair.
=== END OUTPUT ===
```
