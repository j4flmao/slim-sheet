# Role preset: Board / executive pack

**Audience**: board members and executives reading on paper or a tablet. **Decision**: approve,
challenge or ask for more. **Compare**: annual plan and prior year; year-to-date.

Rules beyond the skill: at most 5 KPIs, 3 messages, 2 charts plus one table; no jargon; every
message ends with what it means for the decision.

## Expected columns

`Quarter` · `Segment` · `Revenue (USD m)` · `Revenue plan (USD m)` · `Revenue PY (USD m)` ·
`Net profit (USD m)` · `Net profit plan (USD m)` · `Customers` · `Headcount`

## build_report arguments

```text
title='<Period>: <headline result in one sentence>',
subtitle='<Company> · USD m · YTD <period> vs annual plan and prior year',
kpis=[
  'Revenue=sum:Revenue (USD m)|plan:Revenue plan (USD m):plan|vs:Revenue PY (USD m):PY',
  'Net profit=sum:Net profit (USD m)|plan:Net profit plan (USD m):plan',
  'Net margin=ratio:Net profit (USD m)/Revenue (USD m)',
  'Customers=sum:Customers',
  'Headcount=sum:Headcount',
],
charts=[
  'trend:Quarter:Revenue (USD m)',
  'rank:Segment:Net profit (USD m)',
  'table:Segment:Revenue (USD m)',
],
messages=[
  'We delivered <x>% of the annual revenue plan by <period>; <on track / at risk>.',
  '<Segment> drives <y>% of profit; <risk or opportunity>.',
  'Decision needed: <what the board is asked to approve>.',
],
```
