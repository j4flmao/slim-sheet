# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Changed (breaking, before 1.0)
- Five tools merged to keep the tiers at 8/18/30: `validate_formula` → `evaluate check_only=true`, `split` →
  `convert split_by=...`, `add_validation` + `conditional_format` → `rules kind=...`, `rows_ops` → `range_ops`.
  The CLI answers the old commands with the replacement.
- Tool annotations are explicit: titles, `destructiveHint` (false for working-copy writes, true for `sheet_ops` and
  `range_ops`), `idempotentHint`, `openWorldHint=false`.
- Schema export opens scalar unions (cell values) instead of emitting `anyOf`, for Gemini and OpenAI strict mode.

### Added
- `range_ops` (full): fill (formulas translated, series), copy/move/transpose, sort, replace, insert/delete rows or
  columns, and `where` (SQL condition) to copy or delete matching rows.
- `sheet_layout` (full): freeze, widths, heights, autofilter, hide/unhide, gridlines, zoom, print area, orientation,
  fit to width, protection.
- `annotate` (full): notes, hyperlinks (https/mailto/in-workbook), images (extra `images`), removing objects.
- `what_if` (standard): delta, goal seek and 1-D/2-D sensitivity grids, recomputed offline, never written.
- `anomalies scope=formulas|all`: model audit (error roots, broken references, inconsistent formulas, typed-over
  values, double counting, hardcoded numbers in formulas); `describe_cell trace=precedents|dependents`.
- `format_range merge`, `overview` on a folder, `overview` warnings for macros, external links and connections.
- `pivot native=true`: real, refreshable Excel PivotTables (row field, optional column field, up to 3 values) with
  values pre-filled; `merge how=reconcile` with `compare` and `tolerance`.
- `range_ops values_only=true` (paste values) and Excel-style `move`: references to moved cells follow them.
- Eval suite: 6 tasks for the new tools; tasks a tier cannot serve are skipped.
- Sparklines: `add_chart type=sparkline`, and existing sparklines are now kept through every edit.
- Structured output (`outputSchema` + `structuredContent`) for overview, read, find, query, profile and diff.
- MCP prompts (`dashboard`, `audit_model`, `clean_data`, `format_workbook`) and resources (`slim-sheet://grammar`,
  `slim-sheet://brands`).
- Opt-in audit log (`SLIM_SHEET_AUDIT_LOG`), per-client rate limits and Origin checks for streamable HTTP.
- Release provenance: build attestations, PyPI attestations and an SPDX SBOM.
- Formula guard shared by every writer (`core/guard.py`), now also blocking outbound HYPERLINK/IMAGE; read, find and
  query output marks cell text as file data.
- Design system: brands (`slim`, `office`, `modern`, `mono`, `gov` or a brand `.json` via `SLIM_SHEET_BRAND`),
  profiles (`finance`, `ibcs`, `gov`), named number formats and semantic cell roles, written into every workbook as
  an Office theme plus named cell styles.
- `build_report` (lite): a one-screen branded dashboard from compact strings — KPI tiles with live formulas, charts
  chosen by intent with insight titles, a chart-data sheet and a source footer.
- `style_sheet` (standard): one-call table styling with number-format inference reported per column.
- `format_range` `role`, `create_workbook` `brand`, `add_chart` `type="auto"`, `doughnut`, `highlight`, `brand`.
- Prompt library: `docs/prompts/design.md` (three-line prompts, grammar, brand setup); dashboard and report
  outputs now delegate design to the server.
- Read-only macros (Phase 8): `read sheet='#vba'` lists VBA modules and pages their source (untrusted block),
  `anomalies scope=macros` rates risk (auto-run, shell/COM/network calls, URLs, obfuscation, VBA stomping),
  `convert out_path=<folder>` exports `.bas`/`.cls`, and `overview` warns from the file's content (modules, signed,
  locked, auto-run, XLM sheets, Mark of the Web). In-house MS-OVBA reader on olefile (BSD); macros never run.
- `SLIM_SHEET_MAX_CELLS` (default 3,000,000): files declaring more cells are refused with `too_large` before loading.
- `range_ops op=clear what=contents|formats|rules|all`, `op=split` (text to columns), `copy formats_only=true`.
- `convert` CSV options: `delimiter`, `encoding` (e.g. `cp1252`) and `decimal=','` when reading; output delimiter
  and `encoding='utf-8-sig'` (BOM) when writing.
- `set_table totals=true` (SUBTOTAL totals row) and resize when run again with the same name; `add_chart
  replace=true`; `sheet_layout calc_mode=` and `protect_structure=`; `overview` hidden-content report (very hidden
  sheets, hidden rows/columns/names, `;;;` formats, notes, embedded objects, author fields).
- `write` warns when a spilling formula has non-empty cells in its way, or a value lands in another formula's spill.
- MCP prompt `replace_macro`; 6 macro-phrased eval tasks on a new `macros.xlsm` fixture.
- Prompt library: roles `model-auditor`, `workbook-editor`, `macro-migrator`, `file-reviewer`; outputs
  `audit-report`, `edited-workbook`; `examples.md` with five filled-in sets and one-liners; compact edit and
  macro blocks for small models; the CORE tool map covers the Phase 7–8 tools.
- Prompt recipes: `docs/prompts/recipes/` with about 50 one-paste prompts in 8 packs (finance, sales and
  marketing, HR, operations, data preparation, Excel tasks, reporting, Vietnam).
- `format_range role="input"` leaves cells unlocked, so inputs stay editable when the sheet is protected.
- Prompt library: `outputs/dashboard-pack.md` (overview, growth and segment dashboards plus a chart
  gallery, every chart title checked by a query) and design.md §6–7 (chart chosen from the question, data
  shapes for trends, segments and concentration curves, ready prompts incl. small models).
- `range_ops op=split` warns when it overwrites non-empty cells; `set_table totals=true` no longer sums
  ID-like columns (OrderID, Invoice No, Year, ...).

- Flexible paths without config: absolute paths work in any local folder, relative paths are looked up in
  every `SLIM_SHEET_ROOTS` folder, relative outputs land next to their source, `overview()` with no path
  lists the roots. System/AppData folders, dot-folders, network shares and non-data files stay refused;
  `SLIM_SHEET_STRICT_ROOTS=1` keeps the old roots-only behaviour.

### Fixed
- Formulas using functions newer than Excel 2007 were saved without the `_xlfn.` prefix: `=FILTER(...)` made the
  file impossible to open in Excel, others showed `#NAME?`. Formulas are now stored the way Excel saves them
  (`_xlfn.`, `_xlfn._xlws.`, LET/LAMBDA `_xlpm.` names, all 161 checked in Excel 365) and shown without prefixes.
- Spilling formulas (FILTER, UNIQUE, SEQUENCE, ...) are written as real dynamic arrays (no implicit `@`), and the
  dynamic-array metadata Excel files carry is kept through every edit (openpyxl dropped it, turning spills into
  legacy CSE arrays).
- `recalc` on `.xlsm` restores the original VBA project byte-for-byte if LibreOffice changed it.
- `set_table` over a range with a sheet autofilter produced a file Excel refused to open; the overlapping
  filter now moves into the table, as Excel does.
- `style_sheet`, `format_range`, `build_report` and `add_chart` ignored the brand a workbook was created
  with once no `brand=` was given on a later call; they now detect the workbook's brand from its named
  styles and keep it, instead of silently reverting to `slim`.
- Unknown or misspelled tool parameters (e.g. a typoed `copy_from=`) were silently dropped instead of
  erroring; every tool's schema now sets `additionalProperties: false` and rejects them by name.
- Numbers computed from sums/divisions (e.g. a pivot preview) could show binary noise like
  `611.2259999999999`; cell text now renders floats to 15 significant digits, the same precision Excel
  itself displays.
- `sheet_ops rename` can change only the letter case of a sheet name (`data` -> `Data`), as Excel can.
- `evaluate` replies list recomputed cells as ranges (`Data!Z2:Z131`, at most 10) and print plain numbers
  instead of `np.float64(...)`.
- `style_sheet` formats formula columns from their headers (formulas have no values until Excel
  calculates), and one outlier no longer turns a margin or growth column into plain numbers.
- `anomalies scope=formulas` no longer reports labelled total rows (Total, All..., and the rows under
  them) as inconsistent formulas.
- Freezing panes again on a sheet already frozen on both axes (e.g. `build_report` on a data sheet that
  `sheet_layout` had frozen at `E2`) left duplicate and impossible pane selections, and Excel refused the
  file. Freezing now starts from a clean view, and every save removes such selections, repairing older files.
- Offline evaluation: `INDEX(...)=""` on a blank cell is TRUE and a formula ending on a blank shows 0,
  matching Excel 365.
- Charts written by openpyxl had hidden axes in Excel 365 (`<c:delete>` missing), rounded corners, a border and
  gridlines; every chart is now styled cleanly.
- Autofit sized numbers from their raw digits instead of their formatted display (`0.2723` vs `27.2%`).

## [0.1.0] - 2026-10-20

First release.

### Added
- Lite tier: `overview`, `read`, `find`, `query`, `profile`, `write`, `explain_formula`.
- Standard tier: `clean`, `anomalies`, `merge`, `convert`, `diff`, `undo`, `validate_formula`, `evaluate`, `describe_cell`.
- Full tier: `create_workbook`, `sheet_ops`, `rows_ops`, `name_ops`, `format_range`, `set_table`, `add_validation`,
  `conditional_format`, `add_chart`, `pivot`, `split`, `recalc`.
- Readers for xlsx/xlsm/xltx/xltm/xlsb/xls/ods (calamine), csv/tsv/parquet/json (DuckDB), Apple Numbers (extra
  `numbers`) and password-protected xlsx (extra `crypto`).
- Token-budgeted Markdown/TSV/JSON output with continuation cursors; closed error codes with corrective hints.
- Working-copy writes, atomic saves, a write journal and `undo`.
- Excel-style reference shifting for row/column inserts, deletes and sheet renames.
- 514-function formula catalog; offline evaluation (extra `formula`).
- Streamable HTTP transport with bearer-token auth; `schema` export for OpenAI and Anthropic tool formats.
- CLI generated from the tool signatures, so every tool is also a shell command.
