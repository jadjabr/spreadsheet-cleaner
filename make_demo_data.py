#!/usr/bin/env python3
"""Generate 5 fake, deliberately messy customer lists in ./input for demoing clean.py.

All names, emails and phone numbers are made up. Re-running overwrites the demo files.
"""

import csv
import io
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
INPUT = HERE / "input"

FIRST_NAMES = ["Olivia", "Liam", "Emma", "Noah", "Ava", "Mateo", "Sophia", "Ethan", "Priya", "Lucas",
               "Mia", "Omar", "Chloe", "Diego", "Hannah", "Wei", "Grace", "Andre", "Nora", "Samuel"]
LAST_NAMES = ["Bennett", "Carter", "Nguyen", "Patel", "Alvarez", "Okafor", "Fischer", "Romano", "Kim", "Hughes",
              "Sato", "Delgado", "Morgan", "Haddad", "Reyes", "Novak", "Sullivan", "Tanaka", "Brooks", "Larsen"]
CITIES = [("Boston", "MA", "617"), ("New York", "NY", "212"), ("Chicago", "IL", "312"), ("Austin", "TX", "512"),
          ("Denver", "CO", "303"), ("Seattle", "WA", "206"), ("Portland", "OR", "503"), ("Atlanta", "GA", "404")]
COMPANIES = ["Northwind Supply", "Bluebird Bakery", "Summit Dental", "Harbor Freight Co", "Maple & Finch",
             "Redwood Realty", "Lakeside Fitness", "Copper Kettle Cafe"]


def make_customers(rng, count=40):
    customers = []
    for index, (first, last) in enumerate(rng.sample([(f, l) for f in FIRST_NAMES for l in LAST_NAMES], count)):
        city, state, area = rng.choice(CITIES)
        customers.append({
            "first": first,
            "last": last,
            "email": f"{first}.{last}@example.com".lower(),
            "phone": f"{area}55501{index:02d}",
            "date": date(2019, 1, 1) + timedelta(days=rng.randrange(2400)),
            "city": city,
            "state": state,
            "company": rng.choice(COMPANIES),
        })
    return customers


def messy_text(rng, text):
    if rng.random() < 0.3:
        text = " " * rng.randint(1, 3) + text
    if rng.random() < 0.3:
        text = text + " " * rng.randint(1, 3)
    if " " in text.strip() and rng.random() < 0.3:
        text = text.replace(" ", "  ", 1)
    return text


def messy_email(rng, email):
    roll = rng.random()
    if roll < 0.2:
        email = email.upper()
    elif roll < 0.35:
        email = email.title()
    return messy_text(rng, email)


def messy_phone(rng, digits, allow_number=False):
    a, b, c = digits[:3], digits[3:6], digits[6:]
    styles = [f"({a}) {b}-{c}", f"{a}-{b}-{c}", f"{a}.{b}.{c}", digits, f"+1 {a} {b} {c}", f"1-{a}-{b}-{c}", f"{a} {b} {c}"]
    if allow_number and rng.random() < 0.2:
        return int(digits)
    return rng.choice(styles)


def messy_date(rng, day, allow_datetime=False):
    if allow_datetime and rng.random() < 0.3:
        return datetime(day.year, day.month, day.day)
    styles = [
        day.strftime("%Y-%m-%d"),
        day.strftime("%m/%d/%Y"),
        f"{day.month}/{day.day}/{day.year}",
        day.strftime("%m/%d/%y"),
        day.strftime("%m-%d-%Y"),
        day.strftime("%d-%b-%Y"),
        day.strftime("%B %d, %Y"),
        f"{day.strftime('%b')} {day.day}, {day.year}",
        day.strftime("%Y/%m/%d"),
    ]
    return rng.choice(styles)


def build_rows(rng, customers, fields, size, duplicates, excel=False):
    """Pick customers for one file, then repeat a few so the file has duplicates."""
    picked = rng.sample(customers, size)
    picked += rng.sample(picked, duplicates)
    rng.shuffle(picked)
    makers = {
        "first": lambda c: messy_text(rng, c["first"]),
        "last": lambda c: messy_text(rng, c["last"]),
        "email": lambda c: messy_email(rng, c["email"]),
        "phone": lambda c: messy_phone(rng, c["phone"], allow_number=excel),
        "date": lambda c: messy_date(rng, c["date"], allow_datetime=excel),
        "city": lambda c: messy_text(rng, c["city"]),
        "state": lambda c: c["state"],
        "company": lambda c: messy_text(rng, c["company"]),
    }
    return [[makers[field](customer) for field in fields] for customer in picked]


def csv_bytes(headers, rows):
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer)
    writer.writerow(headers)
    writer.writerows(rows)
    return buffer.getvalue().encode("utf-8")


def xlsx_bytes(headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def build_demo_files():
    """Return the demo files as {file name: file contents}, without writing anything."""
    rng = random.Random(42)
    customers = make_customers(rng)

    # (file name, headers, fields behind those headers, customers, in-file duplicates)
    files = [
        ("customers_east.csv",
         ["First Name", "Last Name", "Email", "Phone", "Signup Date", "City", "State"],
         ["first", "last", "email", "phone", "date", "city", "state"], 16, 3),
        ("customers_west.csv",
         ["first_name", "last_name", "email", "phone", "signup_date", "city", "state"],
         ["first", "last", "email", "phone", "date", "city", "state"], 15, 2),
        ("newsletter_signups.csv",
         ["EMAIL", "FIRST NAME", "LAST NAME", "SIGNUP DATE", "PHONE", "CITY", "STATE"],
         ["email", "first", "last", "date", "phone", "city", "state"], 18, 3),
        ("crm_export.xlsx",
         [" First  Name ", "Last Name ", " Email", "Phone", "Signup Date", "City"],
         ["first", "last", "email", "phone", "date", "city"], 14, 2),
        ("trade_show_leads.xlsx",
         ["FirstName", "LastName", "Email", "Phone", "SignupDate", "City", "State", "Company"],
         ["first", "last", "email", "phone", "date", "city", "state", "company"], 15, 3),
    ]

    contents = {}
    for name, headers, fields, size, duplicates in files:
        excel = name.endswith(".xlsx")
        rows = build_rows(rng, customers, fields, size, duplicates, excel=excel)
        if name in ("customers_east.csv", "crm_export.xlsx"):
            rows.insert(len(rows) // 2, [""] * len(headers) if not excel else [None] * len(headers))
        contents[name] = (xlsx_bytes if excel else csv_bytes)(headers, rows)
    return contents


def main():
    INPUT.mkdir(exist_ok=True)
    for name, content in build_demo_files().items():
        (INPUT / name).write_bytes(content)
        print(f"  Wrote input/{name}")


if __name__ == "__main__":
    main()
