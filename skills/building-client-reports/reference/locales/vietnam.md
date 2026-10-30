# Optional: Vietnamese reports (brand `vn`)

Use only when the audience is Vietnamese. Everything else in the skill applies unchanged.

## What `brand='vn'` changes

- KPI notes in Vietnamese: `+15,6% so với cùng kỳ`, `Hoàn thành 107,3% KH`, `2,4x so với cùng kỳ`.
  The separators follow the reader's Excel settings; the stored formats stay standard.
- Dates `dd/mm/yyyy`; workbook language `vi-VN`.
- Header words are read like their English equivalents for number formats: `Doanh thu`, `LN`,
  `LNST`, `Biên LN`, `Tỷ trọng`, `Tăng trưởng`, `Vốn hóa`, `VCSH`, `P/E`, `ROE`, `%HTKH`, `Mã số`.

## Conventions to follow

| Convention | How |
|---|---|
| Plan vs actual (`TH` vs `KH`) | `|plan:<KH column>:KH` → "Hoàn thành X% KH" |
| Same period last year (`cùng kỳ`) | `|vs:<CK column>:cùng kỳ` |
| Year to date (`lũy kế`) | aggregate the YTD rows; say "Lũy kế 9T/2026" in the subtitle |
| Unit line | subtitle `Đơn vị tính: tỷ đồng` (or `triệu đồng`) |
| Big multiples | ≥2x is written as a multiple automatically ("gấp N lần" reads naturally) |

## Example

```text
build_report(path='kqkd_2026.xlsx', data_sheet='Data', brand='vn',
  title='9T/2026: lợi nhuận vượt 7% kế hoạch năm',
  subtitle='Công ty ABC · Đơn vị tính: tỷ đồng · Lũy kế 9T/2026',
  kpis=['Doanh thu=sum:DT 2026|vs:DT 2025:cùng kỳ|plan:DT KH:KH',
        'LNST=sum:LNST 2026|vs:LNST 2025:cùng kỳ|plan:LNST KH:KH',
        'Biên LN=ratio:LNST 2026/DT 2026'],
  charts=['trend:Tháng:DT 2026', 'rank:Chi nhánh:LNST 2026'])
```

Statutory financial statements (forms B01-DN…B09-DN under Circular 99/2025/TT-BTC, from FY2026)
have a fixed legal layout; build them from the accountant's template, not from a dashboard.
