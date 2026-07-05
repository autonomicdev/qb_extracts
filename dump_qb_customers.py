"""
QuickBooks Desktop ODBC - Dump v_lst_customer to CSV
Outputs all fields from v_lst_customer via the QB File DSN.

Requirements:
  pip install pyodbc

Usage:
  python dump_qb_customers.py --qb-dsn "C:\QB\company.dsn" [--output v_lst_customer.csv]
"""

import csv
import argparse
import pyodbc


def main():
    parser = argparse.ArgumentParser(description="Dump v_lst_customer from QB ODBC to CSV")
    parser.add_argument("--qb-dsn", required=True,
                        help=r"Path to QB ODBC File DSN (e.g. C:\QB\company.dsn)")
    parser.add_argument("--output", default="v_lst_customer.csv", help="Output CSV filename")
    args = parser.parse_args()

    print(f"Connecting via ODBC DSN: {args.qb_dsn}")
    conn = pyodbc.connect(f"FileDSN={args.qb_dsn}", autocommit=True)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM v_lst_customer")
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
    finally:
        conn.close()

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(columns)
        for row in rows:
            writer.writerow([str(v) if v is not None else "" for v in row])

    print(f"Done. {len(rows)} customer(s) written to {args.output}")
    print(f"Columns: {', '.join(columns)}")


if __name__ == "__main__":
    main()
