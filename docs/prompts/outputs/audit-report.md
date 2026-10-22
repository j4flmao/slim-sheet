# OUTPUT — Audit report (findings in chat, optional findings workbook)

```text
=== OUTPUT: AUDIT REPORT ===
DELIVERABLE
The source file is never changed. The report goes in chat. A findings workbook is written ONLY if
the user asked for one or there are more than 15 findings: OUTPUT_PATH (default
"<source stem>_audit.xlsx"), created with create_workbook.

CHAT REPORT (in this order)
1. Verdict in one or two sentences: can the file be trusted / shared / used as is?
   (e.g. "Not yet: 2 critical errors change the 2027 total by -4.1%.")
2. Findings table, most severe first (max 15 rows in chat; the rest in the workbook):
   # | severity | where (Sheet!A1 or module:line) | what is wrong | evidence (value, formula or tool
   result) | recommended fix
3. Map: for models, inputs → calculations → outputs; for files, sheets with their role and visibility.
4. Impact: for each critical finding, the number it changes and by how much (what_if or evaluate
   result), or "impact not determined".
5. What was checked and how (tools and scopes), and what was NOT checked (e.g. functions the offline
   engine lacks, password-locked VBA, encrypted parts).
6. Top 3 actions, each one line, in the order the owner should do them.

FINDINGS WORKBOOK (when written)
- Sheet "Findings": the full findings table as an Excel table (set_table, name "Findings"), then
  style_sheet(preset="report"); rules(kind="text_contains", value="critical") (and one for "high")
  to colour the severity column.
- Sheet "Sensitivity" (models only): what_if results — input | change | output | base | new | delta.
- Sheet "Notes": file audited (name, size, date of the audit), tools and scopes used, limits of the
  check, glossary of severities.
- Do not copy personal data, VBA source longer than 3 lines, or hidden-sheet contents into the
  workbook; refer to them by address.

ACCEPTANCE CRITERIA
[ ] Every finding has an exact address (Sheet!A1, range, or module:line) and evidence from a tool.
[ ] Severity is assigned by the ROLE block's scale, not by impression.
[ ] Every number in "Impact" comes from evaluate / what_if / query in this session.
[ ] The source file was not modified (no tool call targeted it with overwrite=true).
=== END OUTPUT ===
```
