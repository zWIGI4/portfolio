# 05 · KPI dashboard from three systems

![KPI dashboard from three systems](../docs/img/kpi3.png)

**Problem:** a business running recurring classes (swim school, fitness,
courses) keeps payments in one system, sign-ups in a second and attendance in
a third. None of them answers: how much do we really earn, who signed up and
never paid, which group carries the result.

**Solution:** the script joins three exports on email + month and builds one
Excel dashboard. A separate reconciliation sheet lists every mismatch between
the systems, because that's where money usually leaks.

## Input → output
- `input/payments.csv` — 239 payments (Stripe-like format)
- `input/bookings.csv` — 261 bookings (Bookwhen-like)
- `input/attendance.csv` — 980 attendance records (iClass Pro-like)
- `output/kpi_dashboard.xlsx` — 5 sheets:
  - **KPI** — net revenue, active clients, % of bookings paid, attendance
  - **By month** — revenue table with data bars + native chart
  - **By course** — revenue, clients and attendance per group + chart
  - **Reconciliation** — bookings without payment (to chase)
    and payments without booking (to book manually)
  - **Data** — merged raw records as a filterable Excel table

## The hard parts (done)
- three systems with different formats and IDs; the join key is email + month
- revenue counted net of refunds, amounts as numbers
- a reconciliation sheet instead of silently dropping mismatches

## Run it
```bash
python3 make_sample_data.py  # rebuild the fictional exports (or drop in your own)
python3 kpi_dashboard.py     # build output/kpi_dashboard.xlsx
```

## In a real job
Loaders get adapted to your exports (Stripe, Bookwhen, iClass Pro, Przelewy24,
anything CSV/XLSX) and we agree the KPI list. Next month is one run.

Tech: Python, openpyxl (native Excel charts). Data 100% fictional.

---

*🇵🇱 Skrót: płatności + rezerwacje + obecności złączone po e-mail+miesiąc →
jeden dashboard z arkuszem uzgodnienia (rezerwacje bez płatności itd.).
Uruchomienie: `python3 kpi_dashboard.py`. W zleceniu podmieniam wczytywanie pod
Twoje eksporty.*
