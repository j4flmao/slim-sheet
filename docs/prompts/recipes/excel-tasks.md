# Excel task recipes

## 1. Input template with dropdowns

```text
Use the slim-sheet tools only. create_workbook(path=<path>, sheets=["Entry","Lists"]).
On Lists write the allowed values (<list 1>, <list 2>). On Entry write the headers <headers>.
rules(kind="list", values=[...]) for the choice columns down to row <500>, rules(kind="date",
value="<first allowed date>", value2="<last allowed date>") on date columns, rules(kind="decimal", value="0", value2="<max>") on amounts, each with a message.
set_table on the header + first row, sheet_layout(freeze="A2", widths=[...]).
Hide Lists with sheet_ops(op="hide"). Report the rules added per column.
```

## 2. Fix #REF! and #N/A errors

```text
Use the slim-sheet tools only. find(path=<path>, text="#(REF|N/A|VALUE|DIV/0)", regex=true).
anomalies(scope="formulas") to find where each error starts (error_root). For each root:
explain_formula on it and on a healthy neighbour, say what is broken, and propose a fix.
Apply fixes to the working copy with write, evaluate each fixed cell, and run find again to show
the errors are gone.
```

## 3. Add a lookup column

```text
Use the slim-sheet tools only. In <path>, sheet <main>, add a column "<new col>" that looks up
<value col> from sheet <lookup sheet> by <key>. Use =XLOOKUP(<key cell>,<lookup keys>,<lookup values>,"not found")
(or INDEX/MATCH if EXCEL_VERSION is 2016: check with evaluate check_only=true target="2016").
write the first cell, range_ops fill to the last row, evaluate the first and last, and query how
many rows got "not found".
```

## 4. Protect a sheet for others

```text
Use the slim-sheet tools only. In <path>, mark the input cells (<range>) with
format_range(role="input") (input cells are styled and left unlocked; all other cells stay locked), then sheet_layout(sheet=<sheet>, protect=true) and, if needed,
sheet_layout(protect_structure=true). No password is set by the tool; tell me to add one in Excel
(Review > Protect Sheet). Report which ranges remain editable.
```

## 5. Print setup

```text
Use the slim-sheet tools only. For each report sheet in <path>: sheet_layout(landscape=true,
fit_width=true, print_area=<used range>, gridlines=false, freeze=<header row>). Report the print
area set per sheet.
```

## 6. Named ranges for a model

```text
Use the slim-sheet tools only. In <path>, name_ops(op="list") first. Add names for the key inputs
and outputs (<name>=<sheet!cell> ...) with name_ops(op="add"), then show which formulas could use
them (describe_cell trace="dependents" on each input). Do not rewrite formulas unless I ask.
```

## 7. Excel table with totals

```text
Use the slim-sheet tools only. In <path>, sheet <sheet>, make <range> a table named <Name> with a
totals row: set_table(range=<range>, name=<Name>, totals=true). If rows are added later, run
set_table again with the same name and the bigger range to resize it. style_sheet(preset="data").
```

## 8. Convert formulas to values (freeze a report)

```text
Use the slim-sheet tools only. In <path>, copy <range> onto itself as values:
range_ops(op="copy", range=<range>, to=<same top-left cell>, values_only=true) into a new
out_path so the live version is kept. diff both files and confirm only formulas became values.
```
