# HR recipes (privacy-safe)

All HR recipes aggregate. Never paste names, salaries, IDs or contact details of individual people
into chat; report counts and totals, and keep person-level detail inside the workbook.

## 1. Headcount report

```text
Use the slim-sheet tools only. <path> is an employee list (department, location, grade, start
date, end date, employment type). query headcount at <date> (started on or before, not ended)
by department and location, and FTE if a hours column exists. build_report(title="Headcount <date>",
kpis=["Headcount=count", ...], charts=["rank:<department>:count", "share:<employment type>:count"]).
Report totals only, no names.
```

## 2. Attrition

```text
Use the slim-sheet tools only. From <path> compute monthly leavers, average headcount and
attrition rate (leavers / average headcount) for the last 12 months, overall and by department
(query). Write a sheet "Attrition" and add_chart(type="line", title=<the trend in one sentence>).
Departments with fewer than 10 people are grouped as "Other" to protect privacy.
```

## 3. Payroll check before payment

```text
Use the slim-sheet tools only. This month's payroll <path> and last month's <prev path>.
merge(how="reconcile", keys=["<employee id>"], compare=["<gross>", "<net>"], tolerance=<1>).
Report counts only: joiners, leavers, people whose gross changed by more than <10%>, and
negative or zero net pay (query). Write the exceptions to a sheet "Exceptions" in a new file; do not
paste any person's pay in chat.
```

## 4. Leave balance

```text
Use the slim-sheet tools only. <path> has entitlement per employee and a leave log (employee,
start, end, type). Add working-day counts per leave row with =NETWORKDAYS(start,end) (fill down),
then a sheet "Balance": employee id, entitlement, taken (SUMIFS), remaining. rules(kind="less_than",
value="0") on remaining. Report how many people have a negative balance.
```

## 5. Clean an employee list

```text
Use the slim-sheet tools only. clean(path=<path>, ops=["trim","collapse_spaces","unify_null",
"normalize_dates","dedupe"], keys=["<employee id>"]) as a preview first; show the counts of each
change and apply=true to a new file only after the preview looks right. anomalies on the result;
report duplicates and missing ids as counts.
```
