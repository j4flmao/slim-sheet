# Role preset: Market or industry research for a client

**Audience**: a client's management or investment committee. **Decision**: where the market is
going and who wins. **Compare**: prior year, peers, market share.

Rules beyond the skill: cite the source of every dataset in the subtitle or a message; mark
estimates; convert currencies once, at a stated rate, and say which.

## Expected columns

`Year` · `Region` · `Company` · `Sector` · `Revenue (USD bn)` · `Revenue PY (USD bn)` ·
`Net income (USD bn)` · `Market cap (USD bn)` · `Net margin` · `Profit growth`

## build_report arguments

```text
title='<Market> <year>: <top company/region> takes <x>% of profit',
subtitle='<N> companies · USD bn at <year> average FX · source: <source>',
kpis=[
  'Revenue=sum:Revenue (USD bn)|vs:Revenue PY (USD bn):PY',
  'Net income=sum:Net income (USD bn)',
  'Net margin=ratio:Net income (USD bn)/Revenue (USD bn)',
  'Market cap=sum:Market cap (USD bn)',
  'Companies=count:Company',
],
charts=[
  'rank:Region:Net income (USD bn)',
  'share:Sector:Market cap (USD bn)',
  'relationship:Net margin:Profit growth',
  'table:Company:Net income (USD bn)',
],
messages=[
  '<Region> earns <x>% of profit and holds <y>% of market value.',
  'Profit grew <z>% while revenue grew <w>%: margins widened by <v> pts.',
  'Concentration: the top 10 companies hold <u>% of profit.',
],
```
