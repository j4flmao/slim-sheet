# Ready-made prompt sets

Filled-in examples of START + CORE + ROLE + OUTPUT for the most common jobs. Copy the START block,
replace the paths, then paste the listed blocks after it (only the text inside each file's code
block). Each set ends with what a good reply looks like, so you can judge the agent's answer.

For a one-line version of the same job (fine for Claude, GPT-class and Gemini Pro models), see
"One-liners" at the end. For 7–14B local models use [small-models.md](small-models.md) instead.

---

## 1. Audit a financial model before a board meeting

```text
=== START ===
Use the slim-sheet tools only. Follow the CORE PROTOCOL, ROLE and OUTPUT blocks below.
FILES: C:\Finance\budget_2027_model.xlsx
QUESTION / GOAL: Can we trust the 2027 EBITDA on the Summary sheet? Which inputs move it most?
AUDIENCE: CFO
EXCEL_VERSION: 2021
CONSTRAINTS: do not change the model; write findings to C:\Finance\budget_2027_audit.xlsx
=== END START ===
```

Paste after it: [core.md](core.md) → [roles/model-auditor.md](roles/model-auditor.md) →
[outputs/audit-report.md](outputs/audit-report.md)

A good reply has: a verdict sentence with the size of the problem ("EBITDA is overstated by 3.2%");
findings with exact cells (`Opex!F14`) ordered critical → low; the formula text of each wrong cell
next to a healthy neighbour; a sensitivity ranking from `what_if` (e.g. "price +5% → EBITDA +11.8%");
functions that Excel 2021 lacks; and a note of what the offline engine could not recompute.

---

## 2. Replace the macros in a monthly report

```text
=== START ===
Use the slim-sheet tools only. Follow the CORE PROTOCOL, ROLE and OUTPUT blocks below.
FILES: C:\Ops\monthly_report.xlsm
QUESTION / GOAL: Every month someone runs the "PrepareReport" macro. Tell me what it does, whether
it is safe, and produce this month's result without running it.
AUDIENCE: operations team lead
OUTPUT_PATH: C:\Ops\monthly_report_2026-10.xlsm
=== END START ===
```

Paste after it: [core.md](core.md) → [roles/macro-migrator.md](roles/macro-migrator.md) →
[outputs/edited-workbook.md](outputs/edited-workbook.md)

A good reply has: a module/procedure table (which ones run on open); risky lines with
`module:line` (Shell, downloads, file writes) reported first; a plain-words spec of
"PrepareReport" (sort Data by Date, delete rows where Status = "Void", paste values into Summary,
…); one tool call per spec line, each checked (row counts before/after); and a clear list of what
could not be reproduced (e.g. "emails the file to the manager — do this by hand").

---

## 3. A precise edit: add a computed column and a totals row

```text
=== START ===
Use the slim-sheet tools only. Follow the CORE PROTOCOL, ROLE and OUTPUT blocks below.
FILES: C:\Sales\orders_q3.xlsx
QUESTION / GOAL: In the Orders sheet add a column "Margin %" = (Revenue - Cost) / Revenue for every
order, formatted like the other % columns, then make the data an Excel table named Orders with a
totals row. Delete cancelled orders first.
=== END START ===
```

Paste after it: [core.md](core.md) → [roles/workbook-editor.md](roles/workbook-editor.md) →
[outputs/edited-workbook.md](outputs/edited-workbook.md)

A good reply has: the order of operations (delete → new column → fill → copy formats → table);
the row count before and after the delete; the first and last `Margin %` values evaluated; a change
log with exact ranges; a `diff` summary showing nothing else changed; and the working copy path
(`orders_q3.slim.xlsx`).

---

## 4. Check a workbook before sending it to a customer

```text
=== START ===
Use the slim-sheet tools only. Follow the CORE PROTOCOL, ROLE and OUTPUT blocks below.
FILES: C:\Deals\price_offer_acme.xlsx
QUESTION / GOAL: Is this file safe to send to the customer? Look for hidden sheets, our internal
cost or margin data, notes, author names, links to our servers and macros.
AUDIENCE: account manager
CONSTRAINTS: report only, do not change the file
=== END START ===
```

Paste after it: [core.md](core.md) → [roles/file-reviewer.md](roles/file-reviewer.md) →
[outputs/audit-report.md](outputs/audit-report.md)

A good reply has: a verdict ("No — a very hidden sheet 'Cost' holds unit costs for 48 products");
a findings table with BLOCK / FIX BEFORE SHARING / OK TO KEEP; personal or confidential data as
counts and columns, never values; external references with the file names they point to; and the
fix steps, offered on a copy.

---

## 5. Clean a CSV export from an old system, then analyse it

```text
=== START ===
Use the slim-sheet tools only. Follow the CORE PROTOCOL, ROLE and OUTPUT blocks below.
FILES: C:\Exports\erp_invoices_2026.csv
QUESTION / GOAL: The ERP exports with ';' separators, comma decimals and Windows-1252 accents.
Convert it properly, clean it, and tell me which customers have overdue invoices over 60 days.
AUDIENCE: credit control
LANGUAGE of the deliverable: English
CURRENCY / UNITS: EUR
=== END START ===
```

Paste after it: [core.md](core.md) → [roles/data-quality-steward.md](roles/data-quality-steward.md)
→ [outputs/clean-dataset.md](outputs/clean-dataset.md)

A good reply has: the `convert` call with `delimiter=";"`, `encoding="cp1252"`, `decimal=","` and
a check that accents and decimals came through (e.g. "Café", 1234.56); the quality scorecard;
`clean` previewed before applying; and the overdue list as a query result with the SQL shown.

---

## 6. A dashboard pack with many chart types

```text
=== START ===
Use the slim-sheet tools only. Follow the CORE PROTOCOL, ROLE and OUTPUT blocks below.
FILES: D:\gather-data\data\companies_all.csv, D:\gather-data\data\exchanges.csv
QUESTION / GOAL: Where is the profit made, what changed since last year, and how does Vietnam
compare? Build an overview, a growth and a Vietnam dashboard plus a chart gallery.
AUDIENCE: investment committee
OUTPUT_PATH: D:\gather-data\stock_exchanges_dashboards.xlsx
=== END START ===
```

Paste after it: [core.md](core.md) → [roles/general-analyst.md](roles/general-analyst.md) →
[outputs/dashboard-pack.md](outputs/dashboard-pack.md)

A good reply has: three dashboards whose charts use different intents (rank, share, relationship,
compare, trend); a Charts sheet with clustered columns, a pie, a line, an area concentration curve
and sparklines on the main table; every chart title listed with the query that confirmed it (and
any title it had to rewrite); helper sheets hidden; and a note that the file opened in Excel.

---

## One-liners

For strong models a single sentence plus the server's built-in instructions is often enough. Each
names the tools so the agent does not fall back to writing code.

```text
Audit the formulas in C:\Finance\model.xlsx with slim-sheet (anomalies scope=formulas, describe_cell trace=precedents on the key outputs, what_if on the top inputs); change nothing and list findings by severity with cell addresses.
```

```text
Read the macros in C:\Ops\report.xlsm with slim-sheet (read sheet='#vba', anomalies scope='macros'), never run them, explain what each one changes, then redo the main one with range_ops/write on a working copy and verify each step.
```

```text
In C:\Sales\orders.xlsx, sheet Orders: delete rows where Status = 'Cancelled', add column "Margin %" = (Revenue-Cost)/Revenue filled to the last row, then set_table name=Orders totals=true; use slim-sheet tools only and show a diff at the end.
```

```text
Before I send C:\Deals\offer.xlsx outside the company, use slim-sheet overview and find to list hidden content, personal data (counts only), external links, notes, author fields and macros; recommend what to remove; don't change the file.
```

```text
Convert C:\Exports\erp.csv with slim-sheet convert (delimiter=';', encoding='cp1252', decimal=','), check the result with profile, then answer: total amount per customer for invoices overdue more than 60 days.
```

The MCP prompt `replace_macro` (offered by hosts that show MCP prompts, e.g. Claude Desktop's "+"
menu) runs example 2 in one click.
