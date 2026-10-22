# Finance recipes

## 1. Budget vs actual by cost centre

```text
Use the slim-sheet tools only. Files: budget <budget path> (sheet <budget sheet>) and actuals
<actuals path> (sheet <actuals sheet>), both with columns for cost centre, month and amount.
1. overview both files; profile the amount columns.
2. merge(path=<actuals>, other_path=<budget>, keys=["<cost centre>", "<month>"], how="full",
   out_path="<folder>/budget_vs_actual.xlsx") so every cost centre-month appears once.
3. In the merged file add columns Variance (=Actual-Budget) and Variance % (=IFERROR(Variance/Budget,""))
   with write on the first row and range_ops(op="fill") down to the last row.
4. style_sheet(preset="finance"); rules(kind="less_than", value="0") on Variance %.
5. Report: total budget, total actual, the 5 largest overspends (query), and any cost centre that
   exists in only one file.
```

## 2. Month-end variance commentary

```text
Use the slim-sheet tools only. In <path>, sheet <P&L sheet> has accounts in rows and months in
columns. Compare <month> with <previous month> and with the same month last year if present.
query each line's change; list lines whose change is above <threshold, e.g. 10% and 50,000>.
For each, describe_cell(trace="precedents") on the total to show what feeds it. Write a "Variance
notes" sheet: account | this month | last month | change | change % | driver (from the data) |
question for the owner. style_sheet(preset="finance"). Never invent a reason: if the data does
not show the driver, write "to confirm".
```

## 3. Accounts-receivable aging

```text
Use the slim-sheet tools only. <path> has open invoices with customer, invoice date or due date,
and amount outstanding. Aging date: <yyyy-mm-dd>.
1. profile; confirm dates are real dates (clean ops=["normalize_dates"] preview if they are text).
2. query buckets by days past due: current, 1-30, 31-60, 61-90, >90, per customer and in total.
3. Write the result to a sheet "Aging" (customer rows, bucket columns, total), style_sheet
   (preset="finance", total_row=true), rules(kind="data_bar") on the >90 column.
4. add_chart(type="bar", title="<who holds most of the >90 balance>") of the >90 bucket by customer.
5. Report total outstanding, % over 90 days, top 5 customers by >90 balance (all from queries).
```

## 4. 12-month cash-flow forecast

```text
Use the slim-sheet tools only. Create <out path> with create_workbook(sheets=["Inputs","Forecast"]).
Inputs: opening cash <amount>, monthly revenue <amount>, growth per month <x%>, collection rate
<y%> in month 1 and the rest in month 2, fixed costs <amount>, variable costs <z% of revenue>.
Put inputs in Inputs!B column with labels in A, format_range(role="input") on them.
Forecast: 12 month columns; rows Revenue, Cash in, Fixed costs, Variable costs, Net cash flow,
Closing cash, all as formulas referencing Inputs (write the first month, range_ops fill across).
evaluate the closing cash of month 12; what_if(op="grid", inputs=["Inputs!<growth cell>=0%,2%,5%",
"Inputs!<collection cell>=80%,90%,100%"], outputs=["Forecast!<month 12 closing cash>"]).
Report the lowest closing-cash month and the grid.
```

## 5. One-page P&L summary from a general ledger

```text
Use the slim-sheet tools only. <path> is a GL export (date, account, account name, debit, credit).
Map accounts to P&L lines with a sheet "Mapping" (account prefix -> line: Revenue, COGS, Opex,
Other) that you write and I can edit. query the totals per line per month using the mapping
(join on account prefix). Write "P&L" with lines as rows and months as columns plus a Total
column, Gross margin and Net margin rows as formulas. style_sheet(preset="finance"). Check: total
debits = total credits in the source (query) and the P&L net equals revenue minus all costs.
```

## 6. Invoice tie-out between ledger and bank

```text
Use the slim-sheet tools only. Ledger <ledger path>, bank statement <bank path>.
merge(how="reconcile", keys=["<invoice or reference>=<bank reference>"], compare=["<amount>"],
tolerance=<0.01>, out_path="<folder>/tie_out.xlsx"). Report matched, different, ledger-only and
bank-only counts and amounts; list differences above <amount>. Do not change either source file.
```

## 7. Restate figures in another currency

```text
Use the slim-sheet tools only. <path> has amounts in several currencies (column <currency>).
Write a sheet "FX" with currency | rate to <target currency> (I give the rates: <list>).
Add a column "<target> amount" = amount / rate with an INDEX/MATCH on FX, fill down, and a note
(annotate op="note") on the header saying the rate date. query totals per currency before and after
and report them side by side.
```
