# -*- coding: utf-8 -*-
"""Fictional data generator: 3 exports from 3 different systems.

Simulates a sports class school (data is 100% fictional):
- payments.csv    (payment system, Stripe-like format)
- bookings.csv    (sign-up system, Bookwhen-like format)
- attendance.csv  (attendance system, iClass Pro-like format)

The data is deliberately imperfect: some bookings have no payment, some
payments have no booking, and refunds happen. The dashboard has to catch that.
The column names and values stay Polish because they mirror real exports from
Polish systems.
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(77)
INPUT = Path(__file__).parent / "input"

FIRST_NAMES = ["Anna", "Piotr", "Kasia", "Marek", "Ola", "Tomek", "Magda", "Paweł",
               "Ewa", "Michał", "Zosia", "Bartek", "Julia", "Kuba", "Natalia",
               "Adam", "Karolina", "Wojtek", "Maja", "Filip"]
LAST_NAMES = ["Nowak", "Kowalska", "Wiśniewski", "Wójcik", "Kamińska", "Lewandowski",
              "Zielińska", "Szymański", "Dąbrowska", "Kozłowski", "Jankowska",
              "Mazur", "Krawczyk", "Piotrowska", "Grabowski"]

GROUPS = [
    ("Żabki (4-6 lat)", 160, "pon 17:00"),
    ("Delfinki (7-9 lat)", 180, "wt 17:30"),
    ("Rekiny (10-12 lat)", 180, "śr 18:00"),
    ("Młodzież", 200, "czw 19:00"),
    ("Dorośli początkujący", 220, "pt 20:00"),
    ("Dorośli doskonalący", 220, "sob 9:00"),
]

def random_client(i):
    first = random.choice(FIRST_NAMES)
    last = random.choice(LAST_NAMES)
    return f"{first} {last}", f"{first.lower()}.{last.lower().replace('ą','a').replace('ś','s').replace('ż','z').replace('ź','z').replace('ó','o').replace('ł','l').replace('ę','e').replace('ć','c').replace('ń','n')}{i}@example.com"

# 60 clients assigned to groups
clients = []
for i in range(60):
    name, email = random_client(i)
    group = random.choice(GROUPS)
    clients.append({"name": name, "email": email, "group": group[0], "price": group[1]})

MONTHS = [date(2026, m, 1) for m in range(2, 7)]  # February-June 2026

# --- bookings: client x month (monthly pass) ---
bookings = []
booking_id = 3000
for m in MONTHS:
    for c in clients:
        if random.random() < 0.88:  # not everyone books every month
            booking_id += 1
            status = "potwierdzona" if random.random() > 0.05 else "anulowana"
            bookings.append({
                "booking_id": f"BW-{booking_id}",
                "data_startu": (m + timedelta(days=random.randint(0, 5))).isoformat(),
                "kurs": c["group"],
                "klient": c["name"],
                "email": c["email"],
                "cena": c["price"],
                "status": status,
            })

# --- payments: most confirmed bookings have one ---
payments = []
payment_id = 84000
for b in bookings:
    if b["status"] != "potwierdzona":
        continue
    if random.random() < 0.94:  # ~6% of bookings left unpaid
        payment_id += 1
        status = "succeeded"
        amount = b["cena"]
        if random.random() < 0.02:
            status = "refunded"
        payments.append({
            "payment_id": f"py_{payment_id}",
            "data": b["data_startu"],
            "kwota": amount,
            "waluta": "PLN",
            "status": status,
            "email": b["email"],
            "opis": f"Karnet: {b['kurs']}",
        })

# orphan payments (no booking, e.g. a deposit for a camp)
for i in range(8):
    payment_id += 1
    c = random.choice(clients)
    payments.append({
        "payment_id": f"py_{payment_id}",
        "data": (random.choice(MONTHS) + timedelta(days=random.randint(0, 25))).isoformat(),
        "kwota": random.choice([150, 300, 450]),
        "waluta": "PLN",
        "status": "succeeded",
        "email": c["email"],
        "opis": "Wpłata indywidualna",
    })
random.shuffle(payments)

# --- attendance: 4 classes/month per confirmed booking ---
attendance = []
for b in bookings:
    if b["status"] != "potwierdzona":
        continue
    start = date.fromisoformat(b["data_startu"])
    for week in range(4):
        attendance.append({
            "data_zajec": (start + timedelta(days=7 * week)).isoformat(),
            "grupa": b["kurs"],
            "uczestnik": b["klient"],
            "email": b["email"],
            "obecny": "tak" if random.random() < 0.82 else "nie",
        })

def save(name, rows):
    path = INPUT / name
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"  {name}: {len(rows)} rows")

print("Generating fictional exports:")
save("payments.csv", payments)
save("bookings.csv", bookings)
save("attendance.csv", attendance)
