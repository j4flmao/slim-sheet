# OUTPUT — Executive dashboard workbook

```text
=== OUTPUT: EXECUTIVE DASHBOARD ===
DELIVERABLE
One .xlsx at OUTPUT_PATH (default "<first source stem>_dashboard.xlsx") with sheets, in order:
  Dashboard | Dashboard data | KPIs | Analysis | Data | Notes
Never build it inside a source file: convert or copy the cleaned data into the deliverable first.

HOW TO BUILD (the server does the design)
1. Put the cleaned dataset on a sheet named "Data" (one header row, one row per record).
2. Choose 3–6 KPIs from the ROLE metric library and 2–4 charts that answer the main question.
3. One call:
   build_report(path=OUTPUT_PATH, data_sheet="Data", title="<Organisation> — <Subject> <Period>",
                kpis=["Revenue=sum:<col>", "Margin=ratio:<profit col>/<revenue col>", "Orders=count"],
                charts=["trend:<period col>:<measure>", "rank:<dimension>:<measure>",
                        "share:<dimension>:<measure>"], overwrite=true)
   (overwrite=true is allowed here because you created OUTPUT_PATH in this task.)
   Intents: trend (over time), rank (biggest/smallest), compare (few groups), share (mix of a whole),
   relationship (X vs Y). Give a chart a title only when you know a better message than the server's.
4. style_sheet(path, "Data", preset="data"); style_sheet on KPIs and Analysis with preset="report".
5. KPIs sheet: KPI | Definition | Formula/logic | Value | Comparison | Change (one row per tile).
6. Notes sheet: purpose, audience, sources, refresh date, data model, definitions, assumption log,
   data quality log, SQL used.

COMPARISONS
Every KPI needs context. Put the comparison (previous period, target, last year) in the KPIs sheet,
and when it matters, add a KPI tile for it, e.g. "Growth==KPIs!D3/KPIs!E3-1|growth".

ACCEPTANCE CRITERIA (add to CORE Phase 7)
[ ] build_report reply quoted: every KPI value matches an independent query (state both numbers).
[ ] Every chart title states a message; no chart repeats another.
[ ] style_sheet replies read; every wrong inferred format corrected.
[ ] KPIs sheet documents every tile; Notes sheet complete; no "#####" on any sheet.
=== END OUTPUT ===
```

Manual layout (only if `build_report` is not available in your tier): see
[design.md §5](../design.md#5-when-you-build-by-hand-full-tier).
