"""Generate a deliberately messy sales CSV for demoing the report automator."""
import csv
import random
from datetime import date, timedelta

random.seed(42)

PRODUCTS = {"Wireless Mouse": 24.99, "USB-C Hub": 39.99, "Mechanical Keyboard": 89.99,
            "27in Monitor": 229.00, "Webcam HD": 54.50, "Laptop Stand": 32.00}
REGIONS = ["North", "South", "East", "West"]
REPS = ["Alice Chen", "Marcus Lee", "Priya Patel", "Tom Rivera", "Sara Kim"]
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y"]


def messy(value: str) -> str:
    """Randomly add inconsistent casing and stray spaces."""
    r = random.random()
    if r < 0.15:
        return f"  {value.upper()} "
    if r < 0.30:
        return value.lower()
    return value


rows = []
start = date(2026, 1, 1)
for i in range(1, 501):
    d = start + timedelta(days=random.randint(0, 272))
    product = random.choice(list(PRODUCTS))
    qty = random.randint(1, 12)
    price = PRODUCTS[product]
    row = {
        "Order ID": f"ORD-{i:05d}",
        "Order Date": d.strftime(random.choice(DATE_FORMATS)),
        "Region": messy(random.choice(REGIONS)),
        "Sales Rep": messy(random.choice(REPS)),
        "Product": messy(product),
        "Quantity": str(qty),
        "Unit Price": random.choice([f"{price:.2f}", f"${price:,.2f}"]),
    }
    # Inject problems: blanks, bad numbers, duplicates
    if random.random() < 0.03:
        row["Quantity"] = ""
    if random.random() < 0.02:
        row["Unit Price"] = "N/A"
    rows.append(row)
    if random.random() < 0.03:
        rows.append(dict(row))  # duplicate row

with open("sample_data/sales_raw.csv", "w", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=rows[0].keys())
    writer.writeheader()
    writer.writerows(rows)

print(f"Wrote {len(rows)} rows to sample_data/sales_raw.csv")
