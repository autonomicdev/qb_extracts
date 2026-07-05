"""
QuickBooks Desktop - Dump all customer fields to CSV via QBXML COM.
Run this on the QB machine to identify the right customer ID field.

Requirements:
  pip install pywin32

Usage:
  python dump_qb_customers.py [--output v_lst_customer.csv]
"""

import csv
import argparse
import xml.etree.ElementTree as ET

CUSTOMER_QUERY_START = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <CustomerQueryRq requestID="1" iterator="Start">
      <MaxReturned>200</MaxReturned>
      <ActiveStatus>All</ActiveStatus>
      <OwnerID>0</OwnerID>
    </CustomerQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

CUSTOMER_QUERY_CONT = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <CustomerQueryRq requestID="1" iterator="Continue" iteratorID="{iterator_id}">
      <MaxReturned>200</MaxReturned>
    </CustomerQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

CUSTOMER_QUERY_STOP = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <CustomerQueryRq requestID="1" iterator="Stop" iteratorID="{iterator_id}">
    </CustomerQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""


def text(node, tag):
    child = node.find(tag)
    return child.text.strip() if child is not None and child.text else ""


def flatten(node, prefix=""):
    """Recursively flatten an XML element into a dict of tag-path -> text.
    DataExtRet elements are keyed by their DataExtName value."""
    result = {}
    if node.tag == "DataExtRet":
        name = text(node, "DataExtName")
        value = text(node, "DataExtValue")
        if name:
            result[f"DataExt:{name}"] = value
        return result
    for child in node:
        if child.tag == "DataExtRet":
            result.update(flatten(child))
        else:
            key = f"{prefix}/{child.tag}" if prefix else child.tag
            if len(child):
                result.update(flatten(child, key))
            else:
                result[key] = child.text.strip() if child.text else ""
    return result


def fetch_customers():
    import win32com.client
    rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    rp.OpenConnection2("", "QB Customer Dump", 1)
    ticket = rp.BeginSession("", 2)

    all_customers = []

    try:
        def parse(xml_str):
            root = ET.fromstring(xml_str)
            rs = root.find(".//CustomerQueryRs")
            if rs is None:
                raise RuntimeError("No CustomerQueryRs in response")
            status_code = rs.attrib.get("statusCode", "0")
            if status_code not in ("0", "1"):
                raise RuntimeError(f"QB error {status_code}: {rs.attrib.get('statusMessage')}")
            for cust in rs.findall("CustomerRet"):
                all_customers.append(flatten(cust))
            return rs.attrib.get("iteratorID", ""), int(rs.attrib.get("iteratorRemainingCount", "0"))

        iterator_id, remaining = parse(rp.ProcessRequest(ticket, CUSTOMER_QUERY_START))
        print(f"  Fetched {len(all_customers)} customers, {remaining} remaining...")

        while remaining > 0 and iterator_id:
            iterator_id, remaining = parse(
                rp.ProcessRequest(ticket, CUSTOMER_QUERY_CONT.format(iterator_id=iterator_id))
            )
            print(f"  Fetched {len(all_customers)} customers total, {remaining} remaining...")

        if iterator_id:
            rp.ProcessRequest(ticket, CUSTOMER_QUERY_STOP.format(iterator_id=iterator_id))

    finally:
        rp.EndSession(ticket)
        rp.CloseConnection()

    return all_customers


def main():
    parser = argparse.ArgumentParser(description="Dump QB customer fields to CSV via QBXML")
    parser.add_argument("--output", default="v_lst_customer.csv", help="Output CSV filename")
    args = parser.parse_args()

    print("Connecting to QuickBooks on this machine...")
    customers = fetch_customers()

    if not customers:
        print("No customers returned.")
        return

    all_fields = []
    seen = set()
    for c in customers:
        for k in c:
            if k not in seen:
                seen.add(k)
                all_fields.append(k)

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_fields, quoting=csv.QUOTE_ALL, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(customers)

    print(f"Done. {len(customers)} customer(s) written to {args.output}")
    print(f"Fields: {', '.join(all_fields)}")


if __name__ == "__main__":
    main()
