# -*- coding: utf-8 -*-
"""Generator of fictional data for verifying a database of housing developments.

Writes two files that deliberately drift apart:
- internal_database.csv  (our database: 40 developments)
- external_source.csv    (fresh market snapshot: the same developments spelled
  a little differently + changes + 6 new ones + 4 missing)

Injected problems: different spellings of names (dashes, "etap II" vs "etap 2"),
changed prices and dates, typos in district names, new and withdrawn entries.
"""
import csv
import random
from pathlib import Path

random.seed(21)
INPUT = Path(__file__).parent / "input"
INPUT.mkdir(exist_ok=True)

DEVELOPERS = ["Nowe Osiedla SA", "GrupaDom", "Bud-Rem-Invest", "Apartamenty Plus",
              "Zielona Przystań Development", "MegaBud", "Osiedla Jutra"]
CITIES = {
    "Warszawa": ["Mokotów", "Wola", "Ursynów", "Białołęka", "Bemowo"],
    "Kraków": ["Podgórze", "Bronowice", "Prądnik Biały", "Zabłocie"],
    "Wrocław": ["Krzyki", "Fabryczna", "Psie Pole"],
    "Gdańsk": ["Wrzeszcz", "Oliwa", "Jasień"],
    "Poznań": ["Grunwald", "Jeżyce", "Rataje"],
}
ADJECTIVES = ["Zielone", "Słoneczne", "Nowe", "Ciche", "Parkowe", "Miejskie"]
NOUNS = ["Tarasy", "Ogrody", "Wzgórza", "Zacisze", "Aleje", "Bulwary"]

def development_name():
    n = f"{random.choice(ADJECTIVES)} {random.choice(NOUNS)}"
    if random.random() < 0.4:
        n += f" etap {random.choice(['I', 'II', 'III'])}"
    return n

developments = []
used = set()
while len(developments) < 40:
    name = development_name()
    city = random.choice(list(CITIES))
    if (name, city) in used:
        continue
    used.add((name, city))
    developments.append({
        "id": f"INW-{100 + len(developments)}",
        "deweloper": random.choice(DEVELOPERS),
        "nazwa_inwestycji": name,
        "miasto": city,
        "dzielnica": random.choice(CITIES[city]),
        "liczba_lokali": random.randint(24, 320),
        "status": random.choice(["w budowie", "w sprzedaży", "oddana"]),
        "termin_oddania": f"{random.choice([2026, 2027, 2028])}-Q{random.randint(1, 4)}",
        "cena_za_m2": random.randint(9500, 19500),
    })

with open(INPUT / "internal_database.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=list(developments[0].keys()))
    writer.writeheader()
    writer.writerows(developments)

# --- external version: transformations + drift ---
def distort_name(n):
    if "etap I" in n and random.random() < 0.7:
        n = (n.replace("etap III", "etap 3")
              .replace("etap II", "etap 2")
              .replace("etap I", "etap 1"))
    if random.random() < 0.25:
        n = n.replace(" ", "-", 1)
    if random.random() < 0.2:
        n = n.upper()
    if random.random() < 0.15:  # typo: dropped letter (exercises fuzzy matching)
        pos = random.randint(1, len(n) - 2)
        if n[pos].isalpha():
            n = n[:pos] + n[pos + 1:]
    return n

external = []
skipped = random.sample(range(40), 4)  # 4 missing on their side
for i, dev in enumerate(developments):
    if i in skipped:
        continue
    rec = {
        "inwestycja": distort_name(dev["nazwa_inwestycji"]),
        "miasto": dev["miasto"],
        "dzielnica": dev["dzielnica"],
        "lokale": dev["liczba_lokali"],
        "status": dev["status"],
        "oddanie": dev["termin_oddania"],
        "cena_m2": dev["cena_za_m2"],
    }
    draw = random.random()
    if draw < 0.20:      # changed price
        rec["cena_m2"] = int(dev["cena_za_m2"] * random.choice([1.04, 1.07, 0.97]))
    elif draw < 0.32:    # shifted completion date
        year, quarter = dev["termin_oddania"].split("-Q")
        rec["oddanie"] = f"{int(year) + 1}-Q{quarter}" if quarter == "4" else f"{year}-Q{int(quarter) + 1}"
    elif draw < 0.40:    # different status
        rec["status"] = {"w budowie": "w sprzedaży", "w sprzedaży": "oddana",
                         "oddana": "oddana"}[dev["status"]]
    elif draw < 0.46:    # different unit count
        rec["lokale"] = dev["liczba_lokali"] + random.choice([-8, 6, 12])
    external.append(rec)

# 6 new developments we do not have
for _ in range(6):
    city = random.choice(list(CITIES))
    external.append({
        "inwestycja": development_name(),
        "miasto": city,
        "dzielnica": random.choice(CITIES[city]),
        "lokale": random.randint(30, 220),
        "status": random.choice(["w budowie", "w sprzedaży"]),
        "oddanie": f"{random.choice([2027, 2028])}-Q{random.randint(1, 4)}",
        "cena_m2": random.randint(9800, 18900),
    })
random.shuffle(external)

with open(INPUT / "external_source.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=list(external[0].keys()))
    writer.writeheader()
    writer.writerows(external)

print(f"internal_database.csv: {len(developments)} | external_source.csv: {len(external)}")
