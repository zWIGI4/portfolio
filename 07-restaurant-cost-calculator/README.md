# 07 · Restaurant cost calculator (a live Excel tool)

![Restaurant cost calculator](../docs/img/restaurant.png)

**Problem:** the owner knows the rent and the payroll but can't see the whole
picture: what a month of operation costs, how many guests per day break even,
what happens to the result if the average bill goes up by 5 zł.

**Solution:** an Excel workbook where everything is computed by live formulas
from one assumptions sheet. Change a yellow cell — costs, profit and the
break-even point recalculate. This is a tool the business keeps, not a one-off
report.

## What's inside
- `output/restaurant_costs.xlsx` — 3 sheets:
  - **Assumptions** — yellow-filled input cells: rent, utilities, team
    (headcount × rates), food cost %, average bill, guests/day, opening days
  - **Costs** — fixed and variable costs, all formulas over named ranges
  - **Profitability** — revenue, result, margin per bill, break-even
    guests/day + a native chart of result vs guest count

## Why this design
- formulas + named ranges instead of hard-coded numbers: the file lives on
  without a programmer
- no macros, so it also works in Google Sheets and LibreOffice
- the chart is native Excel and recalculates with the assumptions

## Run it
```bash
python3 calculator.py   # builds output/restaurant_costs.xlsx
```
Demo numbers: team of 9, average bill 68 zł, 85 guests/day → result
+17 890 zł/month, break-even at 70 guests/day.

## In a real job
Cost structure gets shaped to your venue (gastro, salon, workshop, clinic),
your categories and rates. You get the file plus a short note on which cells
are yours to edit.

Tech: Python, openpyxl (formulas, named ranges, native chart). Sample numbers.

---

*🇵🇱 Skrót: skoroszyt z żywymi formułami — zmieniasz żółte założenia, Excel
sam przelicza koszty, wynik i próg rentowności (bez makr, działa też w Sheets).
Uruchomienie: `python3 calculator.py`.*
