# Role preset: CFO / finance review

**Audience**: CFO, finance committee, lenders. **Decision**: are we on plan, where is the variance,
what needs action. **Compare**: plan/budget first, prior year second.

## Expected columns (rename to yours)

`Month` · `Business unit` · `Revenue (USD m)` · `Revenue PY (USD m)` · `Revenue budget (USD m)` ·
`Gross profit (USD m)` · `Opex (USD m)` · `Opex budget (USD m)` · `EBITDA (USD m)` ·
`EBITDA budget (USD m)` · `Cash (USD m)`

## build_report arguments

```text
title='<Period>: EBITDA <x>% <ahead of / behind> budget, <driver>',
subtitle='<Company> · USD m · <period> vs budget and prior year',
kpis=[
  'Revenue=sum:Revenue (USD m)|vs:Revenue PY (USD m):PY|plan:Revenue budget (USD m):budget',
  'Gross margin=ratio:Gross profit (USD m)/Revenue (USD m)',
  'Opex=sum:Opex (USD m)|plan:Opex budget (USD m):budget',
  'EBITDA=sum:EBITDA (USD m)|plan:EBITDA budget (USD m):budget',
  'Cash=sum:Cash (USD m)',
],
charts=[
  'trend:Month:EBITDA (USD m)',
  'rank:Business unit:EBITDA (USD m)',
  'trend:Month:Opex (USD m)',
  'table:Business unit:Revenue (USD m)',
],
messages=[
  'EBITDA is <x>% <ahead/behind> budget at <value> USD m.',
  '<Unit> explains <y>% of the gap: <reason>.',
  'Opex is <z>% <under/over> budget; <line> is the largest overrun.',
],
```

Cost lines (Opex, cost, expense) turn red when they rise above plan; slim-sheet inverts the
colour for them automatically.

## Checks before delivery

- Revenue and EBITDA totals equal the P&L sheet (state the cell you compared with).
- `anomalies(scope='formulas')` on the model sheets, `anomalies(scope='client')` on the workbook.
