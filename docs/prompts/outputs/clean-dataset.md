# OUTPUT — Clean, documented dataset

```text
=== OUTPUT: CLEAN DATASET ===
DELIVERABLE
- The cleaned data at OUTPUT_PATH in the requested format (default "<stem>_clean.xlsx"):
    .xlsx  → for people (Data sheet + Data Dictionary + Data Quality sheets)
    .csv   → for other tools (UTF-8; text cells starting with = + - @ are escaped automatically)
    .parquet → for analysts / BI tools (types preserved)
  Use clean(..., apply=true, out_path=...) for cleaning and convert(path, out_path) for format changes.
- For .csv/.parquet deliverables, ALSO produce "<stem>_clean_dictionary.xlsx" with the dictionary and
  quality sheets, because those formats cannot hold documentation.

DATA SHEET RULES (xlsx)
- One table per sheet, header in row 1, no blank rows/columns, no totals rows, no merged cells.
- Column names: unique, descriptive, consistent case; keep original names in the dictionary.
- Types: real dates (yyyy-mm-dd), numbers as numbers, codes/IDs as text when they have leading zeros.
- set_table on the full range; formats per CORE Phase 6; autofit.

DATA DICTIONARY SHEET
column | description | type | unit/currency | example | null % | distinct | source column | notes
(Fill from profile results; descriptions are your best inference — mark them "(inferred)".)

DATA QUALITY SHEET
check | result before | action | result after | rows affected
Include: duplicates, nulls in required columns, type fixes, date fixes, spelling unification,
excluded rows (with reason and total value excluded).

ACCEPTANCE CRITERIA
[ ] Row count in vs out explained exactly (removed duplicates, dropped empty rows, exclusions).
[ ] Totals of every measure before vs after reported; differences explained.
[ ] profile on the output shows the intended types for every column.
[ ] Source files untouched.
=== END OUTPUT ===
```
