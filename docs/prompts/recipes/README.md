# Recipes: ready prompts for everyday spreadsheet jobs

A recipe is one self-contained prompt for one job. Replace the `<...>` parts and paste it; no CORE,
ROLE or OUTPUT block is needed. Use the full blocks ([../README.md](../README.md)) when the job is an
open analysis rather than a known task.

Every recipe:

- names the slim-sheet tools to use, so the agent never falls back to writing code;
- edits a working copy (`<name>.slim.xlsx`), never the source, and ends with a check
  (`read`, `query`, `evaluate` or `diff`);
- asks for the numbers it reports to come from tool results in the same session.

| Pack | Recipes |
|---|---|
| [finance.md](finance.md) | budget vs actual, month-end variance, AR aging, cash-flow forecast, P&L summary, invoice tie-out, FX restatement |
| [sales-marketing.md](sales-marketing.md) | sales dashboard, top customers, commission calculator, pipeline by stage, campaign ROI, price list update |
| [hr.md](hr.md) | headcount report, attrition, payroll check, leave balance, org list clean-up (privacy-safe) |
| [operations.md](operations.md) | inventory reorder list, supplier OTIF, stock aging, order backlog, production plan vs actual |
| [data-prep.md](data-prep.md) | clean a messy export, merge two lists, find duplicates, reconcile two systems, split a file by region, CSV import |
| [excel-tasks.md](excel-tasks.md) | input template with dropdowns, fix #REF!/#N/A, lookup column, protect a sheet, print setup, named ranges, table with totals |
| [reporting.md](reporting.md) | one-page KPI dashboard, monthly management pack, board slides data, period-over-period, executive summary |
| [vietnam.md](vietnam.md) | VND formats, e-invoice reconciliation, Vietnamese names, payroll with insurance, tax declaration data |

Tips:

- Give paths in full (`D:\Data\sales.xlsx`) or relative to your default folder.
- For 7–14B local models, run one recipe step per message, or use [../small-models.md](../small-models.md).
- If a reply contains a number you did not see in a tool result, ask: "show the query for that number".
