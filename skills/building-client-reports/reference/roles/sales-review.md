# Role preset: Sales review (monthly / quarterly business review)

**Audience**: sales leadership and account managers. **Decision**: where to push, which accounts
need help. **Compare**: quota/target and prior period.

## Expected columns

`Month` · `Region` · `Rep` · `Customer` · `Product` · `Revenue (USD k)` · `Quota (USD k)` ·
`Revenue PY (USD k)` · `Deals` · `Order ID`

## build_report arguments

```text
title='<Period>: <x>% of quota, <region> leads',
subtitle='<Company> sales · USD k · <period> vs quota and prior year',
kpis=[
  'Revenue=sum:Revenue (USD k)|plan:Quota (USD k):quota|vs:Revenue PY (USD k):PY',
  'Deals=sum:Deals',
  'Orders=count:Order ID',
  'Avg deal=avg:Revenue (USD k)|number1',
],
charts=[
  'trend:Month:Revenue (USD k)',
  'rank:Region:Revenue (USD k)',
  'share:Product:Revenue (USD k)',
  'table:Customer:Revenue (USD k)',
],
messages=[
  'Revenue reached <x>% of quota; <region> is furthest behind at <y>%.',
  'Top 10 customers bring <z>% of revenue; <customer> grew fastest.',
  '<Product> is now <w>% of the mix, up from <v>%.',
],
```
