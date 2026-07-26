# 08 · B2B quoting system

![B2B quoting system](../docs/img/quoting.png)

**Problem:** every quote is rebuilt by hand in a copied spreadsheet. Someone
retypes prices from the price list, forgets a discount, and the version sent to
the customer no longer matches the one in the folder. A 10-line quote takes half
an hour, and nobody can say what margin it actually leaves.

**Solution:** a live Excel quoting tool. Products come from a dropdown fed by
the catalogue, prices and costs are looked up automatically, and the rep only
edits quantities, discounts and the contract term. Totals, margin, the term
comparison and a print-ready page for the customer recalculate on their own.

## Input → output
- `input/catalogue.csv` — 18 products and services (hardware, subscriptions,
  accessories, services) with billing type, unit price and unit cost
- `input/quote_request.csv` — 10 line items the customer asked for, with
  quantities and per-line discounts
- `output/quotation.xlsx` — 5 sheets:
  - **Quote** — the working sheet: yellow cells for product, quantity, discount,
    contract term and VAT rate; SKU, category, billing, price, cost and margin
    filled by formulas; totals and a summary block
  - **Catalogue** — the price list that feeds the dropdowns and every lookup
  - **Term comparison** — the same basket over 12 / 24 / 36 / 48 months, with
    the selected term marked and a chart
  - **Summary** — contract value, VAT, upfront, recurring, cost, gross margin
    + a revenue / cost / margin chart
  - **Customer quote** — print-ready A4 page, no cost or margin data anywhere

## What it does
- product dropdowns driven by the catalogue, so a quote can only contain items
  that actually exist in the price list
- `INDEX`/`MATCH` pulls SKU, category, billing type, price and cost per line
- separates one-off from recurring items and prices the whole contract as
  `upfront + monthly × term`
- gross margin per line and for the whole contract, computed against unit costs
- term comparison recalculates the entire basket over four contract lengths
- the customer-facing sheet is built from formulas, so it can never drift from
  the internal numbers — and it carries no cost or margin data
- sheets are protected with only the yellow input cells left editable

## Run it
```bash
python3 make_sample_data.py   # rebuild the fictional inputs
python3 quote.py              # build output/quotation.xlsx
```
Demo run: 10 line items, 24-month term → 49 776,96 zł upfront + 2 359,20 zł/month,
contract value 106 397,76 zł net (130 869,24 zł gross), gross margin 33.2%.

## In a real job
The catalogue columns, billing rules and VAT get mapped to your price list, and
the customer page is rebuilt to your letterhead and payment terms. If quotes go
out as PDF, the export step gets added too.

Tech: Python, openpyxl (formulas, named ranges, data validation, native charts,
sheet protection). No macros — the file also opens in LibreOffice and Google
Sheets. All data fictional.

---

*🇵🇱 Skrót: żywy arkusz ofertowy — produkty z listy rozwijanej zasilanej
cennikiem, ceny i koszty podstawiane formułami, a handlowiec zmienia tylko
ilości, rabaty i długość umowy; wartość kontraktu, marża, porównanie okresów
i gotowa do druku strona dla klienta przeliczają się same.
Uruchomienie: `python3 quote.py`.*
