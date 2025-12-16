
"""
create_company_db.py
Creates SQLite DB `company_data.db` with synthetic Orders data (>=100 rows).
"""
import sqlite3, random
from datetime import datetime, timedelta
from pathlib import Path

DB_PATH = "company_data.db"

customers = [
    "Acme Corp","Globex LLC","Stark Industries","Wayne Enterprises","Umbrella Co",
    "Initech","Hooli","Wonka Industries","Soylent Corp","Cyberdyne Systems",
    "Tyrell Corporation","Gringotts Bank","Duff Beer","Pied Piper","Aperture Labs",
    "Nakatomi Trading","Dunder Mifflin","Monsters Inc","Oceanic Airlines","Gekko & Co",
    "Primatech","Massive Dynamic","Vandelay Industries","Bluth Company","Prestige Worldwide"
]

products = [
    "Widget A","Widget B","Gadget Pro","Gadget Mini","Service Plan",
    "Connector","Adapter","Sensor Pack","Premium Support","Maintenance Kit",
    "Battery Pack","Smart Hub","Cloud License","Training Course","Accessory Set"
]

statuses = ["delivered","processing","canceled","returned"]

def create_db(n_rows: int = 300):
    path = Path(DB_PATH)
    if path.exists():
        path.unlink()
    conn = sqlite3.connect(path)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE Orders (
            order_id INTEGER PRIMARY KEY,
            customer_name TEXT NOT NULL,
            product TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            price_per_unit REAL NOT NULL,
            order_date TEXT NOT NULL,
            status TEXT NOT NULL
        );""")
    cur.execute("CREATE INDEX idx_orders_date_status ON Orders(order_date, status);")
    random.seed(42)
    base = datetime.now()
    rows = []
    for i in range(1, n_rows+1):
        c = random.choice(customers)
        p = random.choice(products)
        q = random.randint(1, 20)
        price = round(random.uniform(10, 500) * (1.2 if "Pro" in p or "Premium" in p else 1.0), 2)
        d = (base - timedelta(days=random.randint(0, 119))).date().isoformat()
        st = random.choices(statuses, weights=[0.7,0.15,0.1,0.05])[0]
        rows.append((i, c, p, q, price, d, st))
    cur.executemany("INSERT INTO Orders VALUES (?,?,?,?,?,?,?);", rows)
    conn.commit()
    conn.close()

if __name__ == "__main__":
    create_db()
    print("company_data.db created with Orders table.")
