# Role preset: Marketing performance

**Audience**: CMO, channel owners, agencies. **Decision**: where to move budget. **Compare**:
budget and prior period; efficiency ratios.

## Expected columns

`Week` · `Channel` · `Campaign` · `Spend (USD)` · `Budget (USD)` · `Impressions` · `Clicks` ·
`Leads` · `Revenue (USD)`

## build_report arguments

```text
title='<Period>: <channel> returns <x> per dollar, <channel> under 1',
subtitle='<Company> marketing · USD · <period> vs budget',
kpis=[
  'Spend=sum:Spend (USD)|plan:Budget (USD):budget',
  'Leads=sum:Leads',
  'Cost per lead=ratio:Spend (USD)/Leads|currency',
  'CTR=ratio:Clicks/Impressions|percent',
  'ROAS=ratio:Revenue (USD)/Spend (USD)|multiple',
],
charts=[
  'trend:Week:Leads',
  'rank:Channel:Revenue (USD)',
  'relationship:Spend (USD):Revenue (USD)',
  'table:Campaign:Revenue (USD)',
],
messages=[
  '<Channel> returns <x> USD per USD spent, <y>x the average.',
  '<Campaign> spent <z>% of budget for <w>% of leads.',
  'Moving <amount> from <channel> to <channel> would <expected effect>.',
],
```

Spend is a cost: rising above budget shows red automatically.
