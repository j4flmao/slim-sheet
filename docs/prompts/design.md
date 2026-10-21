# Design — themes, layouts and dashboards with the least description

slim-sheet decides how a workbook looks; the agent only decides **what** to show. This page has
paste-ready prompts, the company brand setup, and the rules the server applies, so anyone can get
company-grade output without writing a style guide into every prompt.

## 1. The shortest prompts that work

**Dashboard from a data sheet** (works on the `lite` tier, also for 7–14B local models):

```text
Use the slim-sheet tools only. Open <path>, find the data sheet with overview, then call
build_report once: title "<subject> <period>", 3–6 KPIs and 2–4 charts that answer "<question>".
KPI items look like "Revenue=sum:Revenue", "Margin=ratio:Profit/Revenue", "Orders=count".
Chart items look like "trend:Month:Revenue", "rank:Region:Revenue", "share:Product:Units".
Then tell me each KPI value and each chart title from the tool reply.
```

**Make an existing table look professional**:

```text
Use the slim-sheet tools only. For every data table in <path>, call style_sheet with
preset="report" (data dumps: preset="data"; financial statements: preset="finance").
Read each reply; if an inferred number format is wrong, fix only that column with
format_range(number_format=<name>). Do not set colours or fonts yourself.
```

**Restyle a workbook in the company brand** (after the admin set `SLIM_SHEET_BRAND`, or name one):

```text
Use the slim-sheet tools only. Apply brand "<slim|office|modern|mono|gov or acme.json>":
style_sheet(..., brand=...) on every table, and rebuild the dashboard with build_report(..., brand=...).
```

## 2. The compact grammar

| Item | Form | Examples |
|---|---|---|
| KPI | `Label=agg:Column\|format` | `Revenue=sum:Revenue`, `Avg order=avg:Total\|currency`, `Top deal=max:Total` |
| KPI ratio | `Label=ratio:Numerator/Denominator` | `Margin=ratio:Profit/Revenue` (percent by default) |
| KPI count | `Label=count` or `Label=count:Column` | `Orders=count`, `Customers=count:CustomerID` |
| KPI formula | `Label==<Excel formula>` | `Target gap==Targets!B2-SUM(Sales!E2:E500)` |
| Chart | `intent:Category:Value[:agg][:Title]` | `trend:Month:Revenue`, `rank:Region:count`, `share:Product:Units:sum:Mix by product` |

Aggregates: `sum` (default), `avg`, `count`, `min`, `max`. Formats: `integer, number1, number2, currency,
percent, growth, multiple, k, m, bn, date, year` — leave them out and the server infers one.

Chart intents (what message the chart carries — the server picks the chart type):

| Intent | Use when the question is… | Server draws |
|---|---|---|
| `trend` | how did it change over time? | line (column when ≤ 3 periods), time order |
| `rank` | who is biggest / smallest? | horizontal bars, sorted, leader highlighted, top 11 + Other |
| `compare` | how do a few groups compare? | sorted columns (bars for long labels or > 8 groups) |
| `share` | what is the mix of a whole? | doughnut for 2–5 parts, otherwise sorted bars |
| `relationship` | does X move with Y? | scatter |
| `auto` | not sure | trend if the categories are periods, otherwise compare |

Every chart gets a title that states its message ("Revenue rose 58% from Jan to Dec", "North leads
Revenue with 35% of the total") unless you give one.

## 3. Company brand (set once, used by every agent)

Built-in brands: `slim` (default, teal accent + colour-blind-safe palette, Calibri), `office`
(classic Office blue), `modern` (Office 2023+, Aptos), `mono` (IBCS-style greys, minus signs),
`gov` (UK government statistics style: no fills, minus signs, `1 Feb 2026` dates).

A company brand is a small JSON file; only `name` is required:

```json
{
  "name": "Acme",
  "base": "slim",
  "primary": "#C8102E",
  "palette": ["#C8102E", "#1F3864", "#7F7F7F", "#F2A900", "#00A3AD", "#6D2077"],
  "heading_font": "Arial",
  "body_font": "Arial",
  "currency": "$",
  "profile": "finance",
  "negatives": "parens"
}
```

Set `SLIM_SHEET_BRAND=C:\Brand\acme.json` in the MCP server config (see [hosts.md](../hosts.md)).
`profile`: `finance` (parentheses for negatives, dash for zero, double-ruled totals), `ibcs`
(minus signs, greys), `gov` (no fills, no zero dash, no frozen panes). The brand is written into each
workbook as a real Office theme, so Excel's own table and chart styles follow it when people keep editing.

## 4. Rules the server applies (so prompts don't have to)

- **Layout** (build_report): title band → KPI row (up to 6 tiles: label, big value, what it measures)
  → 2-column chart grid → source footer; gridlines off; fits one landscape page; chart numbers on a
  `<sheet> data` sheet; rebuilding replaces the sheets.
- **Tables** (style_sheet): one wrapped header row, numeric headers aligned with their numbers, brand
  body font, totals with a rule, widths from the *formatted* values (no `#####`), header frozen.
- **Numbers**: thousands separators, one decimal for rates (`0.0%`), signed growth (`+3.2%`), big
  headline numbers scaled (`7.1m`); ids, codes and years are never given separators.
- **Charts**: no 3-D, no border, no rounded corners, no gridlines on bars, legend only with several
  series, data labels on short bar charts, one accent colour with grey context, axes always visible.
- **Colour means something**: one accent; green/red only for good/bad; finance inputs blue.

## 5. When you build by hand (full tier)

Only when build_report cannot express the layout: write blocks with `write`, then
`style_sheet(range=...)` per table, `format_range(role="title")` for headings, and
`add_chart(type="auto", data_range=..., title="<insight>", highlight="<item>")` per chart. Keep one
screen: at most 6 KPIs and 4 charts on a dashboard sheet; put details on their own sheets.
