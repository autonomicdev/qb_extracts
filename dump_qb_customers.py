"""
QuickBooks Desktop - Dump all v_lst_customer fields to CSV via ODBC.
Run this on the QB machine to identify the right customer ID field.

Requirements:
  pip install pyodbc

Usage:
  python dump_qb_customers.py --dsn "C:\path\to\company.qbw.DSN"
  python dump_qb_customers.py --dsn "C:\Users\Public\Documents\Intuit\QuickBooks\Company Files\autonomic.qbw.DSN"
"""

import csv
import argparse


def main():
    parser = argparse.ArgumentParser(description="Dump v_lst_customer fields to CSV via QB ODBC")
    parser.add_argument(
        "--dsn",
        default=r"C:\Users\Public\Documents\Intuit\QuickBooks\Company Files\autonomic.qbw.DSN",
        help="Path to the QB ODBC File DSN",
    )
    parser.add_argument("--output", default="v_lst_customer.csv", help="Output CSV filename")
    args = parser.parse_args()

    import pyodbc
    print(f"Connecting via FileDSN: {args.dsn}")
    conn = pyodbc.connect(f"FileDSN={args.dsn}", autocommit=True)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM v_lst_customer")
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
    finally:
        conn.close()

    print(f"Columns: {', '.join(columns)}")
    print(f"Rows: {len(rows)}")

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(columns)
        writer.writerows(rows)

    print(f"Done. Written to {args.output}")


if __name__ == "__main__":
    main()
