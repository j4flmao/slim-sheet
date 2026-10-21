# slim-sheet prompt library

Reusable, data-agnostic prompts for doing professional analysis with the slim-sheet MCP server in any
agent (Claude, Codex, OpenCode, Cursor, Gemini CLI, Amp, LM Studio, Ollama hosts…). They never assume
what your files contain: the agent discovers the structure, documents its assumptions and verifies
every number before it reports.

## How it works: START + CORE + ROLE + OUTPUT

A full prompt is four blocks pasted in this order:

| Block | What it decides | File |
|---|---|---|
| START | your files, question, audience, period, output path (all optional) | template below |
| CORE | the method: discover → profile → quality → define → analyse → build → verify → report | [core.md](core.md) |
| ROLE | the business lens: default questions, how to map columns, metric library, pitfalls | [roles/](roles/) |
| OUTPUT | the deliverable: sheets, layout, formats, acceptance criteria | [outputs/](outputs/) |

Copy only the text inside each file's code block.

**In a hurry?** [design.md](design.md) has three-line prompts that produce a branded dashboard or a
professionally formatted table in one tool call (`build_report`, `style_sheet`), the compact KPI/chart
grammar, and how a company sets its brand once with `SLIM_SHEET_BRAND`.

## START block (fill what you know, delete the rest)

```text
=== START ===
Use the slim-sheet tools only. Follow the CORE PROTOCOL, ROLE and OUTPUT blocks below.
FILES: <paths, a folder, or "all spreadsheets in <folder>">
QUESTION / GOAL: <what you need answered or built — or leave blank to use the role defaults>
AUDIENCE: <e.g. CEO, sales managers, finance team>
PERIOD: <e.g. 2026-01 to 2026-09 vs same period 2025 — or blank to use the full data range>
OUTPUT_PATH: <e.g. C:\Reports\q3_dashboard.xlsx — or blank for the default>
LANGUAGE of the deliverable: <English | Vietnamese | ...>
EXCEL_VERSION: <365 | 2021 | 2019 | 2016>
CURRENCY / UNITS: <e.g. VND, values in millions — or blank to detect>
CONSTRAINTS: <e.g. exclude internal customers, fiscal year starts in April, hide salaries>
=== END START ===
```

## Roles

| Role | Use for | Default output |
|---|---|---|
| [general-analyst](roles/general-analyst.md) | unknown or mixed data; first look at a new dataset | report-workbook |
| [sales-analyst](roles/sales-analyst.md) | revenue, orders, customers, targets, pricing, discounts | dashboard |
| [finance-analyst](roles/finance-analyst.md) | P&L, budget vs actual, variance, cost control, run-rate | report-workbook |
| [marketing-analyst](roles/marketing-analyst.md) | spend, funnel, CPA/CAC, ROAS, campaign performance | dashboard |
| [customer-analyst](roles/customer-analyst.md) | retention, cohorts, churn, RFM, LTV, product affinity | report-workbook |
| [operations-analyst](roles/operations-analyst.md) | OTIF, lead times, backlog, inventory health, suppliers | dashboard |
| [hr-analyst](roles/hr-analyst.md) | headcount, attrition, tenure, absence, pay structure (privacy-safe) | dashboard |
| [reconciliation-analyst](roles/reconciliation-analyst.md) | match two sources, explain every difference, aging | reconciliation-report |
| [data-quality-steward](roles/data-quality-steward.md) | audit and clean data before anyone analyses it | clean-dataset |

## Outputs

| Output | Deliverable |
|---|---|
| [dashboard](outputs/dashboard.md) | one-screen executive dashboard built with `build_report`: KPI tiles, 2–4 charts, KPIs and Notes sheets |
| [report-workbook](outputs/report-workbook.md) | Summary of findings + one sheet per question + Data, Data Quality, Notes |
| [clean-dataset](outputs/clean-dataset.md) | cleaned xlsx/csv/parquet + data dictionary + quality log |
| [reconciliation-report](outputs/reconciliation-report.md) | matched / differences / one-sided items + waterfall of the difference |
| [quick-answer](outputs/quick-answer.md) | answer in chat only, nothing written |

Any role can be combined with any output (e.g. finance-analyst + dashboard for a CFO view, or
data-quality-steward + quick-answer for a fast health check).

## Example (sales dashboard)

Paste, in order:
1. the START block with `FILES: C:\Data\sales_2026.xlsx, C:\Data\customers.csv`, `AUDIENCE: CEO`
2. [core.md](core.md)
3. [roles/sales-analyst.md](roles/sales-analyst.md)
4. [outputs/dashboard.md](outputs/dashboard.md)

The agent will inventory the files, find the grain and keys itself, clean (preview first), define the
KPIs, build `sales_2026_dashboard.xlsx`, verify each KPI tile against an independent query, and
report findings, assumptions and checks.

## Choosing the tier and model

- The server's default tier is `full` (30 tools); all patterns need it except `quick-answer`.
- Frontier models (Claude, GPT-class, Gemini Pro): use the full blocks as they are.
- Local 7–14B models: use [small-models.md](small-models.md) — a compact single block with the same
  method — and one question at a time.

## Tips for reliable results

- Answer the agent's numbered open questions by number ("1: use net revenue, 2: fiscal year = April").
- To refine, keep the same blocks and change only START (e.g. a new QUESTION); the method stays stable.
- When the result looks wrong, ask the agent to show the SQL and the tie-out for that number.
- Nothing is written to your source files; deliverables have their own paths and every edit can be
  undone (`undo`).
