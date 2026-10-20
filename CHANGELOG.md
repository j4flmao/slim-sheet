# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

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
