# Client-grade report standards (checklist)

What leading standards agree on, and who enforces it in slim-sheet. **Auto** = slim-sheet does it
without being asked. **Check** = `anomalies(scope='client')` reports it. **You** = the agent decides.

## Contents

1. Message and titles
2. Numbers and units
3. Charts
4. Accessibility
5. Workbook hygiene
6. Print and PDF
7. Where the sources disagree

## 1. Message and titles

| Rule | Who | Source |
|---|---|---|
| Each chart and the report title state one message as a sentence with a number | Auto (chart titles), You (report title, messages) | Consulting action titles; UK Analysis Function charts guide |
| A grey second line says what is measured, in which unit, by what | Auto (`<value> by <category>`), You (subtitle: who · unit · period) | IBCS / ISO 24896 title rule (who/what/when) |
| 3–5 key messages under the KPIs | You (`messages=`) | Minto pyramid; IBCS "SAY" |
| Source line at the bottom | Auto (footer) | UK Analysis Function; consulting source lines |

## 2. Numbers and units

| Rule | Who | Source |
|---|---|---|
| Every amount has a unit (header or one line above the table) | Auto (subtitle units), Check (`missing_unit`) | FAST 3.05; ICAEW Code |
| Thousands separators; decimals from the data's precision; rates 0.0%; growth +0.0% | Auto | Banking conventions; FAST number resolution |
| Negatives with parentheses or minus, never colour alone | Auto, Check (`red_only_negative`) | Excel Accessibility Checker |
| KPI compared with prior, plan or target; ≥2x shown as a multiple | Auto with `\|vs:` / `\|plan:` | IBCS variance rules; management reporting practice |
| Ratios change in percentage points | Auto | IBCS; banking |

## 3. Charts

| Rule | Who | Source |
|---|---|---|
| Time on the horizontal axis (line/columns); categories as sorted horizontal bars | Auto (intent) | IBCS / ISO 24896 rule 2; FT Visual Vocabulary |
| Bars and areas start at zero; no 3-D; no dual axes | Auto, Check (`chart_3d`, `dual_axis`) | FT Visual Vocabulary; UK Analysis Function |
| ≤4 lines, ≤2 paired bar series, ≤4 doughnut parts (else sorted bars) | Auto, warnings | UK Analysis Function; Datawrapper |
| Labels on bars instead of a value axis; longest bars labelled inside | Auto | IBCS label rule |
| One accent colour with grey context; green/red only for good/bad | Auto | IBCS semantic colour; newsroom style |

## 4. Accessibility (WCAG 2.2 AA)

| Rule | Who | Source |
|---|---|---|
| Chart marks ≥3:1 against the background; readable colours first | Auto (palette order), warning for faint brand colours | WCAG 1.4.11 |
| Text ≥4.5:1 (3:1 for ≥18 pt or ≥14 pt bold); faint brand colours fall back to ink | Auto | WCAG 1.4.3 |
| Alt text on every chart | Auto (from the title), Check (`chart_no_alt_text`) | WCAG 1.1.1; Excel Accessibility Checker |
| Document title and language set | Auto (build_report), Check (`no_title`) | Excel PDF/UA guidance |
| No merged cells inside tables; meaningful sheet names | Check | Excel Accessibility Checker; UK Analysis Function |

## 5. Workbook hygiene

| Rule | Who | Source |
|---|---|---|
| Calculation automatic, iterative calculation off | Check | ICAEW Code |
| No links to other files | Check (`external_link`) | ICAEW; FAST |
| Nothing hidden | Check (`hidden_sheet`) | ICAEW; FAST |
| Inputs blue on a fill, formulas black, links green | Auto (roles in style_sheet/format_range) | Banking conventions; ICAEW (fill, not font alone) |
| No hard-coded numbers inside formulas; consistent formulas per row | Check (`anomalies scope='formulas'`) | FAST 3.04; ICAEW |

## 6. Print and PDF

| Rule | Who | Source |
|---|---|---|
| Dashboard: one landscape page, title and "Page N of M" footer | Auto | Common report practice |
| Tables: fit to width, header row repeated; very wide tables across pages at full size | Auto | Common report practice |
| Export with Excel's Save as PDF (tags, alt text) rather than Print to PDF | You | Microsoft Excel PDF accessibility |

## 7. Where the sources disagree (and the default)

- **Colour**: IBCS uses colour only for good/bad; newsrooms use one brand accent. Default: accent +
  grey; `brand='mono'` for strict IBCS.
- **Titles**: IBCS wants an identifying title, consultants a message. Default: both (message, then
  grey what/by line).
- **Pies**: limits range from 4 to 8 slices. Default: doughnut for 2–4 parts, sorted bars above.
- **Cell colours**: banks (blue input/black formula/green link) vs FAST (red export/blue import).
  Default: banking.
