# Data cleaning change report

Input: `clients_raw.csv` — **19 rows**
Output: `clients_clean.xlsx` — **14 clean records**

## What was fixed
- Junk rows removed (missing name / invalid e-mail): **2**
- Duplicates removed (same e-mail + date + amount): **3**
- Dates normalized to ISO `YYYY-MM-DD` (from 5 notations, incl. spelled-out Polish)
- Amounts converted to numbers ("1 234,50 zł", "2200,00", "999.99" → 1234.5, 2200.0, 999.99)
- Phones unified to `+48XXXXXXXXX`
- Tax IDs to 10 digits, no separators
- E-mails lowercased, whitespace trimmed; cities capitalized

## Remaining gaps for the client to fill
- Missing date: **0** | missing amount: **2** | missing phone: **0** | missing tax ID: **6**

## Quick KPIs from the clean data
- Order total: **30 825,98 zł**
- Average order: **2 568,83 zł** (from 12 records with an amount)
