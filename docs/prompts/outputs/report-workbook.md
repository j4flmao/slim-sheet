# OUTPUT — Analytical report workbook

```text
=== OUTPUT: ANALYTICAL REPORT WORKBOOK ===
DELIVERABLE
One .xlsx at OUTPUT_PATH (default "<first source stem>_report.xlsx") with sheets, in order:
  Summary | <one sheet per question or analysis area> | Data | Data Quality | Notes

SHEETS
- Summary:
    B1 title (bold, size 16), B2 scope line (sources, period, refresh date).
    B4 "Key findings": numbered findings, one per row: claim | number | comparison | source sheet.
    Below: the metrics table (metric | value | comparison | change | definition).
- Analysis sheets (name them by question, e.g. "Trend", "By Region", "Customers", "Variance"):
    each sheet starts with a one-line question in B1 (bold) and the answer in B2 (one sentence with
    the number). Tables start at B4, one empty row between tables, each table with a bold title row
    above it. Add at most one chart per sheet, anchored to the right of the first table
    (two columns after its last column), titled with the insight.
- Data: cleaned dataset(s) as Excel tables (set_table).
- Data Quality: issue | where | rows | action | impact on totals.
- Notes: data model, metric definitions, assumption log, SQL used per finding.

TABLE STANDARDS
- Columns ordered: dimension(s) → measures → ratios → change. Sorted meaningfully.
- After writing each table: style_sheet(path, sheet, range=<table>, preset="report") — it styles the
  header and total row, infers number formats (ratios 0.0%, changes signed) and sizes columns. Fix
  only the columns whose inferred format is wrong (format_range number_format=<name>).
- Titles: format_range(role="title") for B1, role="section" for table titles.
- Charts: add_chart(type="auto", title="<insight>", highlight="<item to emphasise>").
- rules: kind=data_bar on ranked measures; kind=less_than / greater_than on changes.
- Prefer formulas that reference Data for totals that should stay live; static query/pivot outputs
  are labelled "(snapshot <date>)" in the table title.

ACCEPTANCE CRITERIA
[ ] Every finding on Summary links to a sheet that contains its supporting table.
[ ] Every analysis sheet has its question and one-sentence answer at the top.
[ ] Tie-outs: totals of each breakdown table equal the headline totals (state in Notes).
[ ] No "#####", no formula errors, all formats explicit.
=== END OUTPUT ===
```
