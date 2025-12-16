
"""
Robot: Aggregation of sales by customers (variant 6-10)
- Connects to SQLite database `company_data.db`
- Aggregates total purchases per customer for the last 30 days (only delivered orders)
- Writes top-5 to table `top_customers` (creates or replaces table)
"""
import sqlite3
from datetime import datetime

DB_PATH = "company_data.db" 

DDL = """
CREATE TABLE IF NOT EXISTS top_customers (
    customer_name TEXT PRIMARY KEY,
    total_amount REAL NOT NULL,
    orders_count INTEGER NOT NULL,
    first_order_date TEXT NOT NULL,
    last_order_date TEXT NOT NULL,
    generated_at TEXT NOT NULL
);
"""

UPSERT = """
INSERT INTO top_customers (customer_name, total_amount, orders_count, first_order_date, last_order_date, generated_at)
VALUES (?, ?, ?, ?, ?, ?)
ON CONFLICT(customer_name) DO UPDATE SET
    total_amount=excluded.total_amount,
    orders_count=excluded.orders_count,
    first_order_date=excluded.first_order_date,
    last_order_date=excluded.last_order_date,
    generated_at=excluded.generated_at;
"""

def run():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(DDL)
    conn.commit()

    query = """
        SELECT
            customer_name,
            ROUND(SUM(quantity * price_per_unit), 2) AS total_amount,
            COUNT(*) AS orders_count,
            MIN(order_date) AS first_order_date,
            MAX(order_date) AS last_order_date
        FROM Orders
        WHERE status='delivered'
          AND order_date >= date('now', '-30 day')
        GROUP BY customer_name
        ORDER BY total_amount DESC
        LIMIT 5;
    """
    rows = cur.execute(query).fetchall()

    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    for customer_name, total_amount, orders_count, first_dt, last_dt in rows:
        cur.execute(UPSERT, (customer_name, total_amount, orders_count, first_dt, last_dt, now))

    conn.commit()

    return rows

if __name__ == "__main__":
    rows = run()
    for r in rows:
        print(r)