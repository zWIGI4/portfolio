# 03 · Sales dashboard from a raw CSV

![Sales dashboard](../docs/img/dashboard.png)

**Problem:** you have a transaction export (hundreds of rows) and zero insight:
how much revenue, which month performs, which category carries the result.
Manual pivot tables, rebuilt every month.

**Solution:** one script turns the raw CSV into a formatted Excel dashboard.
Swap the input file, run once — you get a fresh report.

## Input → output
- `input/sales.csv` — 260 fictional transactions (2025)
- `output/dashboard.xlsx` — 4 sheets:
  - **KPI** — six tiles: revenue, orders, average order, top category,
    best month, channel share
  - **By month** — Excel table with data bars + native bar chart
  - **By category** — table with % share + chart
  - **Data** — raw records as a filterable Excel table

## Why it's convenient
- charts are **native Excel charts** — the client edits them like their own
- numbers are numbers, everything sums and filters
- next month = swap the CSV and run once

## Run it
```bash
python3 dashboard.py
```

## In a real job
Works on any tabular export (shop, ERP, bank, CRM) — columns get mapped to your
format and we agree which KPIs and charts sit on top.

Tech: Python, openpyxl (formatting + native Excel charts).

---

*🇵🇱 Skrót: surowy eksport transakcji → dashboard Excel (kafle KPI, wykresy
natywne, tabele z filtrami). Uruchomienie: `python3 dashboard.py`. W zleceniu
mapuję kolumny pod Twój format i ustalamy KPI.*
