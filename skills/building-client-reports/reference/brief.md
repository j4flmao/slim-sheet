# The report brief

Ten one-line fields, always in this order. A brief with these filled produces the same report from
a frontier model and from a 7B local model, because every field maps to a tool argument.

| # | Field | What to write | Maps to |
|---|---|---|---|
| 1 | AUDIENCE | who reads it and what they already know | tone of messages, which KPIs |
| 2 | DECISION | the decision the report supports | the messages |
| 3 | PERIOD / GRAIN | period covered and the time grain | subtitle, `trend` charts |
| 4 | COMPARE | prior period, plan/budget, target, peers | `\|vs:` and `\|plan:` segments |
| 5 | UNITS | currency and scale (USD m, EUR k, units) | headers, subtitle |
| 6 | KPIS | 3–6 measures with their aggregation | `kpis=[...]` |
| 7 | CHARTS | 2–4 questions as intents | `charts=[...]` |
| 8 | MESSAGES | 3 findings, each with one number and one comparison | `messages=[...]` |
| 9 | CHECKS | totals that must reconcile, caveats to show | `anomalies`, `messages` |
| 10 | OUTPUT | file, sheet name, brand | `out_path`, `sheet`, `brand` |

## Template (copy, fill, paste)

```text
Use the slim-sheet tools only. Follow the building-client-reports skill.
FILE: <path>
AUDIENCE: <e.g. CFO and regional heads>
DECISION: <e.g. where to cut Q4 spend>
PERIOD/GRAIN: <e.g. FY2026 YTD, monthly>
COMPARE: <e.g. prior year column 'Revenue 2025'; budget column 'Budget'>
UNITS: <e.g. USD m>
KPIS: <e.g. Revenue, Gross margin, Orders, Cost>
CHARTS: <e.g. trend of revenue by month; rank regions by revenue; top 10 customers table>
MESSAGES: <write them, or "draft 3 from the data and show them to me first">
CHECKS: <e.g. revenue total equals the P&L sheet B12>
OUTPUT: <e.g. sheet 'Dashboard', brand 'slim', file report.xlsx>
```

## Filled example (global retail)

```text
FILE: sales_2026.xlsx
AUDIENCE: CEO and country managers, monthly business review
DECISION: which countries get extra Q4 marketing budget
PERIOD/GRAIN: Jan–Sep 2026, monthly
COMPARE: prior year 'Revenue 2025 (USD k)'; plan 'Plan 2026 (USD k)'
UNITS: USD k
KPIS: Revenue, Gross margin %, Orders, Average order value
CHARTS: trend of revenue by month; rank countries by revenue; share of revenue by channel;
        top 10 products table
MESSAGES: draft 3 from the data, show them to me first
CHECKS: total revenue equals 'P&L'!C8
OUTPUT: sheet 'MBR Sep 2026', brand 'slim'
```

Resulting call:

```text
build_report(path='sales_2026.xlsx', data_sheet='Orders', sheet='MBR Sep 2026',
  title='Q3 2026: revenue 4% ahead of plan, Europe behind',
  subtitle='Acme Retail · USD k · Jan–Sep 2026 vs 2025 and plan',
  kpis=['Revenue=sum:Revenue 2026 (USD k)|vs:Revenue 2025 (USD k):2025|plan:Plan 2026 (USD k):plan',
        'Gross margin=ratio:Gross profit (USD k)/Revenue 2026 (USD k)',
        'Orders=count:Order ID', 'Avg order=avg:Revenue 2026 (USD k)|number1'],
  charts=['trend:Month:Revenue 2026 (USD k)', 'rank:Country:Revenue 2026 (USD k)',
          'share:Channel:Revenue 2026 (USD k)', 'table:Product:Revenue 2026 (USD k)'],
  messages=['Revenue is 4.1% ahead of plan and 9.8% above 2025.',
            'Germany and France are 12% behind plan; all other countries are ahead.',
            'Online is now 46% of revenue, up from 38% in 2025.'])
```
