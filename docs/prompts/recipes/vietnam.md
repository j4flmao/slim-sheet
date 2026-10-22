# Vietnam recipes

Prompts stay in English (models follow them more reliably); deliverable text can be Vietnamese:
add "Write sheet titles and notes in Vietnamese." to any recipe.

## 1. VND number formats

```text
Use the slim-sheet tools only. In <path>, format VND amounts without decimals and with thousands
separators: format_range(number_format="#,##0 \"₫\"") on <ranges>, or "integer" for plain numbers.
Amounts in millions: number_format="m" and say "triệu đồng" in the header. Percentages: "percent".
Check with describe_cell on one cell of each column.
```

## 2. E-invoice (hóa đơn điện tử) reconciliation

```text
Use the slim-sheet tools only. E-invoice list exported from the tax portal <einvoice path> and the
purchase ledger <ledger path>. Keys: invoice symbol + invoice number + seller tax code
(build one key column in each with =<symbol>&"-"&<number>&"-"&<tax code>, fill down).
merge(how="reconcile", keys=["<key>"], compare=["<amount before VAT>","<VAT amount>"], tolerance=1).
Report: invoices in the ledger but not on the portal (risk for VAT deduction), on the portal but not
booked, and amount differences. Sources unchanged.
```

## 3. Vietnamese full names

```text
Use the slim-sheet tools only. In <path>, column <name col> holds full names (family name first,
e.g. "Nguyễn Văn An"). Insert two empty columns to its right (range_ops op="insert"), then write
Family name =LEFT(<c>,FIND(" ",<c>)-1) and Given name =TRIM(RIGHT(SUBSTITUTE(<c>," ",REPT(" ",50)),50)),
fill down, evaluate two rows. Middle names stay in the full-name column. clean(ops=["trim",
"collapse_spaces"]) first if names have double spaces.
```

## 4. Payroll with social insurance

```text
Use the slim-sheet tools only. Create <out path> with sheets Inputs and Payroll. Inputs: insurance
rates (employee: social 8%, health 1.5%, unemployment 1%; employer: as given by me), the salary cap
for insurance, personal deduction <amount> and dependant deduction <amount> per person, and the
PIT brackets as a table. Payroll per employee id: gross, insured salary =MIN(gross, cap),
employee insurance, taxable income, PIT (lookup on the brackets), net. Formulas only, filled down;
style_sheet(preset="finance", total_row=true). evaluate one employee by hand-checkable numbers.
Do not paste individual salaries in chat; report totals.
```

## 5. Data for a tax declaration

```text
Use the slim-sheet tools only. From the sales ledger <path>, query output VAT by rate (0%, 5%, 8%,
10%, not subject) for <month or quarter>: number of invoices, amount before VAT, VAT. Write a
sheet "VAT summary" in the layout of the declaration lines I give (<lines>), style_sheet, and check
that the sum by rate equals the ledger total (query both). Do not submit anything; this is data
preparation only.
```
