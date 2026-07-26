# 02 · Web scraping: website → ready dataset

![Web scraping to dataset](../docs/img/scraper.png)

**Problem:** the data you need (listings, products, prices, catalogs) lives on
a website, and you need it in Excel for analysis, price comparison or import.
Retyping it by hand costs hours and produces typos.

**Solution:** a script walks the catalog pages, extracts the fields and returns
a tidy dataset in XLSX and CSV — prices as numbers, ready to work with.

## Demo source
`books.toscrape.com` — a public sandbox built **for** practicing scraping (no
ToS restrictions). In a real job the source and field list are yours.

## Input → output
- input: catalog pages (3 by default, `--pages N` for more)
- `output/products.xlsx` and `output/products.csv` — 60 records
- fields: `title, price (number), currency, rating 1–5, availability, url`

## Good practices baked in
- `User-Agent` header and a delay between requests (polite scraping)
- forced UTF-8 and regex-based price parsing — no "Â£51.77" mojibake
- pagination end and HTTP error handling
- styled Excel table: branded header, zebra stripes, frozen header row

## Run it
```bash
python3 scraper.py            # 3 pages
python3 scraper.py --pages 5  # more
```

## In a real job
You give me the URL and the field list — you get the dataset plus a rerunnable
script. I only scrape sources that allow it (robots.txt / ToS checked first).

Tech: Python, requests, BeautifulSoup, openpyxl.

---

*🇵🇱 Skrót: strona WWW → czysty dataset XLSX/CSV, grzeczny scraping (UA,
opóźnienia, robots.txt). Uruchomienie: `python3 scraper.py`. W zleceniu
podmieniam źródło i pola pod Ciebie.*
