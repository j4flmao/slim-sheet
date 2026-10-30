# Role preset: Operations / supply chain

**Audience**: COO, plant and logistics managers. **Decision**: where service or cost is slipping.
**Compare**: target levels and prior period.

## Expected columns

`Week` · `Site` · `Product` · `Units shipped` · `Units ordered` · `On-time orders` · `Orders` ·
`Defects` · `Cost (USD k)` · `Cost target (USD k)` · `Lead time (days)`

## build_report arguments

```text
title='<Period>: on-time delivery <x>%, <site> below target',
subtitle='<Company> operations · <period> vs target and prior period',
kpis=[
  'Fill rate=ratio:Units shipped/Units ordered|percent',
  'On time=ratio:On-time orders/Orders|percent',
  'Defect rate=ratio:Defects/Units shipped|percent',
  'Cost=sum:Cost (USD k)|plan:Cost target (USD k):target',
  'Lead time=avg:Lead time (days)|number1',
],
charts=[
  'trend:Week:Units shipped',
  'rank:Site:Defects',
  'trend:Week:Lead time (days):avg',
  'table:Product:Units shipped',
],
messages=[
  'On-time delivery is <x>%, <y> pts below the <z>% target.',
  '<Site> produces <w>% of defects on <v>% of volume.',
  'Lead time rose <n> days since <period>; <cause>.',
],
```
