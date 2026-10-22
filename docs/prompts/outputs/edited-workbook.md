# OUTPUT — Edited workbook (working copy + change log)

```text
=== OUTPUT: EDITED WORKBOOK ===
DELIVERABLE
The same workbook with the requested change, written to a working copy:
- default: "<stem>.slim.<ext>" next to the source (what every edit tool does without out_path);
- or OUTPUT_PATH when the user gives one (pass out_path on the FIRST edit, then keep editing that
  file with overwrite=true — only that file, never the source).
Macro-enabled files stay macro-enabled (.xlsm in, .xlsm out) and keep their VBA project unchanged.

RULES
- Change only what was asked. Keep sheet order, names, formats, tables, rules, charts and hidden
  states as they were, unless changing them is the request.
- Keep formulas live: when the source uses formulas for a column, the new rows/columns use formulas
  too (fill), not pasted numbers — unless the user asked for values.
- Match the existing look: copy formats from a neighbouring column with
  range_ops(op="copy", formats_only=true) instead of inventing colours or number formats.
- One logical change per call; read each reply, including its warnings.

CHANGE LOG (in chat, always)
| # | sheet | range | change | rows/cells affected | tool |
Then: the diff summary (diff source vs working copy), and one line per warning a tool returned.

VERIFICATION (in chat, always)
- Before/after evidence for each change: a count, a total, or the first and last changed values
  (read / query / evaluate).
- "Untouched" check: diff shows no changes outside the ranges in the change log. If it does, explain
  each one (e.g. formulas that Excel-style moves rewrote) or undo.
- If the change cannot be done exactly with tools, say what is missing and stop instead of
  approximating.

ACCEPTANCE CRITERIA
[ ] The working copy opens with the same sheets plus/minus only the requested ones.
[ ] Every changed cell is inside a range listed in the change log.
[ ] Formulas written are evaluated without errors (evaluate on first and last cell).
[ ] The source file is unchanged.
=== END OUTPUT ===
```
