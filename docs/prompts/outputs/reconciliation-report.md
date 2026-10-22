# OUTPUT — Reconciliation report

```text
=== OUTPUT: RECONCILIATION REPORT ===
DELIVERABLE
One .xlsx at OUTPUT_PATH (default "<stem A>_vs_<stem B>_reconciliation.xlsx") with sheets:
  Summary | Matched | Amount Differences | Only in <A> | Only in <B> | Duplicates | Notes

SHEETS
- Summary:
    totals and counts per source; difference; match statistics table
    (class | records A | records B | amount A | amount B);
    waterfall table: starting difference → timing → amount → duplicates → missing → mapping →
    unexplained remainder; key normalisation rules applied; tolerance used.
- Matched: key, amount A, amount B (equal within tolerance).
- Amount Differences: key, amount A, amount B, difference, reason class.
- Only in A / Only in B: full records from merge(how="anti") each way, with a reason class column
  (timing, missing, mapping, unknown).
- Duplicates: keys duplicated in either source with counts.
- Notes: sources, extraction dates, matching rules, tolerance, assumptions, SQL used.
Use merge(..., out_path=...) to produce the detail sheets' data, then write/format them.

FORMATTING
- Amount columns #,##0.00, differences +#,##0.00;-#,##0.00;0.00 with light_red on non-zero.
- Autofit every sheet; set_table on each detail sheet so the user can filter.

ACCEPTANCE CRITERIA
[ ] Counts: matched + differences + only-A = records A (after duplicate handling); same for B.
[ ] Waterfall sums to the starting difference; unexplained remainder stated.
[ ] Every unmatched record has a reason class (or "unknown").
=== END OUTPUT ===
```
