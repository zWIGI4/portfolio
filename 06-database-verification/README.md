# 06 · Database verification against an external source

![Database verification](../docs/img/verification.png)

**Problem:** you keep a working database (developments, products, partners) and
receive a fresh market snapshot. The same items exist in both files but spelled
differently: dashes, letter case, "etap II" vs "etap 2", typos. Comparing them
by hand takes days and still misses things.

**Solution:** the script pairs records despite spelling differences and returns
an Excel report: what matches, what differs (field by field), what to add, what
to explain. Nothing is deleted automatically — the report points, a human
decides.

## Input → output
- `input/internal_database.csv` — 40 developments (internal database)
- `input/external_source.csv` — 42 items (market snapshot, deliberately
  divergent)
- `output/verification_report.xlsx` — 5 sheets:
  - **Summary** — KPI tiles: pairs, matches, discrepancies, gaps
  - **Discrepancies** — one row per difference: field, our value, source
    value, match method
  - **To add** — items present in the source, missing from ours
  - **To explain** — our items absent from the source
  - **Matching** — pairs with no differences

## How matching works
- normalization: lowercase, Polish diacritics stripped, dashes as spaces,
  Roman numerals ("etap II" = "etap 2")
- exact match on normalized name + city first
- then fuzzy match (difflib, 85% threshold) within the same city — catches typos
- every pair records its match method, so the result is auditable

## Run it
```bash
python3 make_sample_data.py   # rebuild the fictional inputs
python3 verify.py             # build output/verification_report.xlsx
```

## In a real job
The match key and compared fields get adapted to your data (tax ID, product
code, address). Useful for base enrichment, migrations and periodic data QA.

Tech: Python (standard library + difflib), openpyxl. Data 100% fictional.

---

*🇵🇱 Skrót: nasza baza vs zewnętrzny zrzut — parowanie odporne na literówki
i inne zapisy, raport różnic pole po polu, nic nie znika bez decyzji.
Uruchomienie: `python3 verify.py`.*
