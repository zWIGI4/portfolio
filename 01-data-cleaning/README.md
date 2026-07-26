# 01 · Data cleaning: messy CSV → analysis-ready Excel

![Data cleaning](../docs/img/cleaning.png)

**Problem:** an export merged from several sources is a mess — dates in five
formats (including spelled-out Polish ones), amounts stored as text like
"1 234,50 zł", duplicated transactions, junk rows, phone numbers and tax IDs
written every possible way. You can't sum it, filter it, or import it anywhere.

**Solution:** one script tidies the whole file and returns two deliverables — a
clean, sorted register and a change report listing exactly what was fixed or
removed, so you don't have to take the result on faith.

## Input → output
- `input/clients_raw.csv` — 19 rows of typical chaos (fictional)
- `output/clients_clean.xlsx` — 14 clean records as a styled Excel table
- `output/change_report.md` — what was fixed, plus quick KPIs from clean data

## What it does
- dates → ISO `YYYY-MM-DD` (5 different notations, including spelled-out)
- amounts → real numbers that sum ("1 234,50 zł" → 1234.5)
- phones → `+48XXXXXXXXX`, tax IDs → 10 digits, no separators
- emails lowercased, cities capitalized, whitespace trimmed
- duplicate transactions removed (email + date + amount) and junk rows dropped,
  each removal counted in the report

## Run it
```bash
python3 clean.py
```

## In a real job
You send your file (or an anonymized sample) plus two sentences about the
output you need. Cleaning rules get matched to your data; you receive the clean
file, the change report and the script, so you can rerun it on future exports.

Tech: Python, pandas, openpyxl.

---

*🇵🇱 Skrót: brudny CSV (5 formatów dat, kwoty tekstem, duplikaty) → czysty
rejestr Excel + raport zmian. Uruchomienie: `python3 clean.py`. W zleceniu
reguły czyszczenia dopasowuję do Twoich danych.*
