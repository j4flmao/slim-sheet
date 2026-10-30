# Role preset: People / HR

**Audience**: CHRO, leadership team. **Decision**: retention and hiring priorities. **Compare**:
plan headcount and prior period. Privacy: report groups of 5 or more people only; never names.

## Expected columns

`Month` · `Department` · `Location` · `Headcount` · `Headcount plan` · `Leavers` · `Hires` ·
`Salary cost (USD k)` · `Salary budget (USD k)` · `Training hours`

## build_report arguments

```text
title='<Period>: attrition <x>%, <department> highest',
subtitle='<Company> people · <period> vs plan · groups of 5 or more',
kpis=[
  'Headcount=sum:Headcount|plan:Headcount plan:plan',
  'Attrition=ratio:Leavers/Headcount|percent',
  'Hires=sum:Hires',
  'Salary cost=sum:Salary cost (USD k)|plan:Salary budget (USD k):budget',
],
charts=[
  'trend:Month:Headcount',
  'rank:Department:Leavers',
  'share:Location:Headcount',
  'table:Department:Headcount',
],
messages=[
  'Attrition is <x>% annualised, <y> pts above last year.',
  '<Department> loses <z>% of its people; exit reasons point to <reason>.',
  'Salary cost is <w>% of budget with headcount at <v>% of plan.',
],
```
