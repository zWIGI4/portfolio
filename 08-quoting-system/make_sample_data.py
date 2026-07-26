# -*- coding: utf-8 -*-
"""Build the fictional input files for the quoting system.

Writes:
  input/catalogue.csv      18 products and services with prices and unit costs
  input/quote_request.csv  what one customer asked for (line items + discounts)

Every company name, price and product is made up. Product names are Polish
on purpose — the demo shows a Polish-market catalogue, while all column
headers stay in English.
"""
import csv
from pathlib import Path

IN = Path(__file__).parent / "input"
IN.mkdir(exist_ok=True)

# sku, product, category, billing, unit_price, unit_cost
CATALOGUE = [
    ("TEL-SMART-128", "Smartfon biznesowy 128 GB", "Hardware", "one-off", 2149.00, 1720.00),
    ("TEL-SMART-256", "Smartfon biznesowy 256 GB", "Hardware", "one-off", 2699.00, 2180.00),
    ("TEL-SMART-LT", "Smartfon podstawowy 64 GB", "Hardware", "one-off", 1099.00, 860.00),
    ("TEL-TAB-11", "Tablet 11 cali LTE", "Hardware", "one-off", 1849.00, 1490.00),
    ("TEL-LAP-14", "Laptop biznesowy 14 cali", "Hardware", "one-off", 4590.00, 3820.00),
    ("TEL-ROUT-5G", "Router 5G do biura", "Hardware", "one-off", 899.00, 690.00),
    ("SUB-VOICE-M", "Abonament głosowy bez limitu", "Subscription", "monthly", 49.00, 28.00),
    ("SUB-DATA-50", "Pakiet danych 50 GB", "Subscription", "monthly", 39.00, 21.00),
    ("SUB-DATA-UNL", "Pakiet danych bez limitu", "Subscription", "monthly", 79.00, 44.00),
    ("SUB-MDM", "Zarządzanie flotą urządzeń (MDM)", "Subscription", "monthly", 24.00, 11.00),
    ("ACC-CASE", "Etui wzmocnione", "Accessory", "one-off", 129.00, 62.00),
    ("ACC-CHRG", "Ładowarka sieciowa 65 W", "Accessory", "one-off", 189.00, 98.00),
    ("ACC-DOCK", "Stacja dokująca USB-C", "Accessory", "one-off", 749.00, 520.00),
    ("ACC-HEAD", "Słuchawki z redukcją szumów", "Accessory", "one-off", 649.00, 430.00),
    ("SRV-SETUP", "Wdrożenie i konfiguracja urządzeń", "Service", "one-off", 1200.00, 640.00),
    ("SRV-TRAIN", "Szkolenie dla zespołu (dzień)", "Service", "one-off", 2400.00, 1300.00),
    ("SRV-SUPPORT", "Wsparcie techniczne SLA 8h", "Service", "monthly", 89.00, 42.00),
    ("SRV-INSUR", "Ubezpieczenie sprzętu", "Service", "monthly", 19.00, 9.50),
]

# product, quantity, discount (as a fraction of the unit price)
QUOTE_REQUEST = [
    ("Smartfon biznesowy 256 GB", 12, 0.08),
    ("Etui wzmocnione", 12, 0.00),
    ("Ładowarka sieciowa 65 W", 12, 0.10),
    ("Abonament głosowy bez limitu", 12, 0.05),
    ("Pakiet danych 50 GB", 12, 0.05),
    ("Zarządzanie flotą urządzeń (MDM)", 12, 0.00),
    ("Laptop biznesowy 14 cali", 3, 0.06),
    ("Stacja dokująca USB-C", 3, 0.00),
    ("Wdrożenie i konfiguracja urządzeń", 1, 0.00),
    ("Wsparcie techniczne SLA 8h", 12, 0.00),
]

with open(IN / "catalogue.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["sku", "product", "category", "billing", "unit_price", "unit_cost"])
    w.writerows(CATALOGUE)

with open(IN / "quote_request.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["product", "quantity", "discount"])
    w.writerows(QUOTE_REQUEST)

print(f"OK -> input/catalogue.csv ({len(CATALOGUE)} products)")
print(f"OK -> input/quote_request.csv ({len(QUOTE_REQUEST)} line items)")
