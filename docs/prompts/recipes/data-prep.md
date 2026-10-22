# Data preparation recipes

## 1. Clean a messy export

```text
Use the slim-sheet tools only. overview <path>, then profile and anomalies on <sheet>.
clean(ops=["trim","collapse_spaces","remove_nonprintable","unify_null","numbers_from_text",
"normalize_dates","drop_empty_rows","drop_empty_cols"]) as a preview; check the date order
(day-first or month-first) in the samples; then apply=true with out_path=<clean path>.
profile the result and report row counts and column types before and after.
```

## 2. Merge two lists

```text
Use the slim-sheet tools only. Left <left path> (sheet <s1>), right <right path> (sheet <s2>).
Find the key columns with overview/profile (they may be named differently: keys=["<left>=<right>"]).
merge(how="left", out_path=<out path>) and report: rows in, rows out, unmatched left rows
(merge how="anti"), and keys duplicated on the right (query HAVING count(*) > 1).
```

## 3. Find and remove duplicates

```text
Use the slim-sheet tools only. In <path>, find duplicate rows by <key columns> with query
(GROUP BY ... HAVING count(*) > 1) and show 5 examples. Then clean(ops=["dedupe"],
keys=["<key>"]) as a preview, and apply=true to a new file. Report how many rows were removed and
which copy was kept (first occurrence).
```

## 4. Reconcile two systems

```text
Use the slim-sheet tools only. System A <a path>, system B <b path>, key <key>, compare
<columns>. merge(how="reconcile", compare=[...], tolerance=<t>, out_path=<out path>).
Report matched / different / only in A / only in B (counts and amounts), and the 10 largest
differences. Sources stay unchanged.
```

## 5. Split a file by a column

```text
Use the slim-sheet tools only. convert(path=<path>, sheet=<sheet>, split_by="<column>",
mode="files", out_path="<folder>") to write one file per <column> value (mode="sheets" for one
workbook with a sheet each). Report the number of rows per part and check they add up to the
source (query).
```

## 6. Import a CSV from another system

```text
Use the slim-sheet tools only. <csv path> comes from <system> with ';' separators, ',' decimals
and Windows-1252 text. convert(path=<csv>, out_path=<xlsx>, delimiter=";", decimal=",",
encoding="cp1252"). profile the result: numbers must be numeric and accents readable
("é", "ơ"). If not, try encoding="utf-8" or "cp1258". Report the column types.
```

## 7. Combine monthly files

```text
Use the slim-sheet tools only. overview <folder> lists the monthly files. For each file query
its rows with a literal month column ('<yyyy-mm>' AS month) and append them with write into
<out path>, sheet "All" (header once). query the row count per month in the result and compare with
each source's count.
```
