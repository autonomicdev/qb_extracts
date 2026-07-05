"""
QuickBooks Desktop - Dump all customer fields to CSV via QBXML COM.
Run this on the QB machine to identify the right customer ID field.

Requirements:
  pip install pywin32

Usage:
  python dump_qb_customers.py [--output v_lst_customer.csv]
  python dump_qb_customers.py --qb-host 192.168.0.81 [--qb-port 5000]
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


class LocalBackend:
    def __init__(self):
        import win32com.client
        self.rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        self.rp.OpenConnection2("", "QB Customer Dump", 1)
        self.ticket = self.rp.BeginSession("", 2)

    def request(self, xml):
        return self.rp.ProcessRequest(self.ticket, xml)

    def close(self):
        self.rp.EndSession(self.ticket)
        self.rp.CloseConnection()


class RemoteBackend:
    def __init__(self, host, port):
        import requests
        self._requests = requests
        self._url = f"http://{host}:{port}/qbxml"

    def request(self, xml):
        resp = self._requests.post(
            self._url,
            data=xml.encode("utf-8"),
            headers={"Content-Type": "text/xml; charset=utf-8"},
            timeout=120,
        )
        resp.raise_for_status()
        return resp.text

    def close(self):
        pass


def flatten(node, prefix=""):
    """Recursively flatten an XML element into a dict of tag-path -> text."""
    result = {}
    for child in node:
        key = f"{prefix}{child.tag}" if not prefix else f"{prefix}/{child.tag}"
        if len(child):
            result.update(flatten(child, key))
        else:
            result[key] = child.text.strip() if child.text else ""
    return result


def fetch_customers(backend):
    all_customers = []

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

    iterator_id, remaining = parse(backend.request(CUSTOMER_QUERY_START))
    print(f"  Fetched {len(all_customers)} customers, {remaining} remaining...")

    while remaining > 0 and iterator_id:
        iterator_id, remaining = parse(
            backend.request(CUSTOMER_QUERY_CONT.format(iterator_id=iterator_id))
        )
        print(f"  Fetched {len(all_customers)} customers total, {remaining} remaining...")

    if iterator_id:
        backend.request(CUSTOMER_QUERY_STOP.format(iterator_id=iterator_id))

    return all_customers


def main():
    parser = argparse.ArgumentParser(description="Dump QB customer fields to CSV via QBXML")
    parser.add_argument("--output", default="v_lst_customer.csv", help="Output CSV filename")
    parser.add_argument("--qb-host", default=None,
                        help="IP/hostname of QB machine running qb_server.py (omit to run locally)")
    parser.add_argument("--qb-port", type=int, default=5000)
    args = parser.parse_args()

    if args.qb_host:
        backend = RemoteBackend(args.qb_host, args.qb_port)
        print(f"Remote mode: {args.qb_host}:{args.qb_port}")
    else:
        backend = LocalBackend()
        print("Local mode: connecting to QuickBooks on this machine")

    try:
        customers = fetch_customers(backend)
    finally:
        backend.close()

    if not customers:
        print("No customers returned.")
        return

    # Collect all field names seen across all customers
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
