"""Create the synthetic input files in data/raw/.

Three monthly order exports from a fictional coffee roaster. They are messy on
purpose, the way real exports from different systems usually are:

- July: webshop CSV, comma separated, ISO dates, dot decimals.
- August: accounting CSV in Norwegian, semicolon separated, dd.mm.yyyy dates,
  prices like "1 234,50", different column names.
- September: an Excel file with yet another set of column names.

On top of that: duplicated orders, product names with odd spacing and case,
missing quantities, a few unreadable dates and prices.

All names and numbers are made up. Fixed seed, so the files are reproducible.
Usage: python tools/make_sample_data.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"

PRODUCTS = {
    "Ethiopia Yirgacheffe 250g": 169,
    "Colombia Huila 250g": 149,
    "Kenya Nyeri 250g": 179,
    "House Espresso 1kg": 429,
    "Decaf Brazil 250g": 139,
    "Filter Papers (100)": 59,
}
# How the same products show up when typed by hand or exported by other systems.
MESSY_NAMES = {
    "Ethiopia Yirgacheffe 250g": ["ethiopia yirgacheffe 250 g", "  Ethiopia Yirgacheffe 250g", "ETHIOPIA YIRGACHEFFE 250G"],
    "Colombia Huila 250g": ["Colombia  Huila 250g", "colombia huila 250 g"],
    "House Espresso 1kg": ["House espresso 1 kg", "House Espresso 1kg "],
    "Filter Papers (100)": ["filter papers (100)"],
}
CITIES = ["Oslo", "Oslo", "Oslo", "Bergen", "Trondheim", "Stavanger", "Drammen", "Tromsø"]
CHANNELS = ["Webshop", "Webshop", "Webshop", "Café", "Wholesale"]


def orders_for(month, rnd, start_id):
    first = date(2026, month, 1)
    days = (date(2026, month + 1, 1) - first).days
    rows = []
    for i in range(rnd.randint(140, 170)):
        product = rnd.choice(list(PRODUCTS))
        name = product
        if product in MESSY_NAMES and rnd.random() < 0.3:
            name = rnd.choice(MESSY_NAMES[product])
        channel = rnd.choice(CHANNELS)
        qty = rnd.randint(6, 24) if channel == "Wholesale" else rnd.randint(1, 3)
        rows.append({
            "id": f"HL-{start_id + i}",
            "date": first + timedelta(days=rnd.randrange(days)),
            "city": rnd.choice(CITIES),
            "product": name,
            "qty": qty,
            "price": PRODUCTS[product],
            "channel": channel,
        })
    # Duplicates: the same order exported twice.
    rows += [dict(r) for r in rnd.sample(rows, 6)]
    # Missing quantity.
    for r in rnd.sample(rows, 3):
        r["qty"] = None
    rnd.shuffle(rows)
    return rows


def nok(value):
    """Format like Norwegian accounting software: 1 234,50"""
    whole, frac = f"{value:,.2f}".split(".")
    return whole.replace(",", " ") + "," + frac


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    rnd = random.Random(2026)

    july = orders_for(7, rnd, 10001)
    with open(RAW / "webshop_orders_2026-07.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["order_id", "date", "customer_city", "product", "quantity", "unit_price", "channel"])
        for r in july:
            w.writerow([r["id"], r["date"].isoformat(), r["city"], r["product"], r["qty"] or "", f"{r['price']:.2f}", r["channel"]])

    august = orders_for(8, rnd, 20001)
    august[5]["date"] = "31.02.2026"  # impossible date
    with open(RAW / "regnskap_ordre_2026-08.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Ordrenr", "Dato", "By", "Produkt", "Antall", "Pris", "Kanal"])
        for r in august:
            d = r["date"] if isinstance(r["date"], str) else r["date"].strftime("%d.%m.%Y")
            w.writerow([r["id"], d, r["city"], r["product"], r["qty"] or "", nok(r["price"]), r["channel"]])

    september = orders_for(9, rnd, 30001)
    september[7]["price"] = "n/a"  # unreadable price
    df = pd.DataFrame([{
        "Order ID": r["id"],
        "Order date": pd.Timestamp(r["date"]),
        "City": r["city"],
        "Item": r["product"],
        "Qty": r["qty"],
        "Price NOK": r["price"],
        "Sales channel": r["channel"],
    } for r in september])
    df.to_excel(RAW / "orders_september_2026.xlsx", index=False)

    for path in sorted(RAW.iterdir()):
        print(path.name)


if __name__ == "__main__":
    main()
