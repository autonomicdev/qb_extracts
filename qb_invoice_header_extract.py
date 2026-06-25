"""
QuickBooks Desktop - Invoice / Credit Memo Header Extract
Outputs CSV matching the qbInvoiceHeader format.

Local mode (run directly on the QB machine):
  python qb_invoice_header_extract.py [--output qbInvoiceHeader.csv] [--from-date YYYY-MM-DD] [--to-date YYYY-MM-DD]
  Requirements: Windows + QuickBooks Desktop open, pip install pywin32

Remote mode (run from any machine on the LAN):
  python qb_invoice_header_extract.py --qb-host 192.168.0.81 [--qb-port 5000] [options]
  Requirements: pip install requests
  The QB machine must be running qb_server.py.
"""

import csv
import argparse
import os

COLUMNS = [
    "customer_id",
    "transaction_date",
    "doc_num_h",
    "amount",
    "stage",
    "transaction_id",
    "open_balance",
    "po_num",
]

# Custom DataExt field name for the opportunity stage on invoice/credit memo headers
STAGE_FIELD_NAME = "Stage"


def _templates(rq_name):
    start = f"""\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <{rq_name} requestID="1" iterator="Start">
      <MaxReturned>100</MaxReturned>
      {{date_filter}}
      <OwnerID>0</OwnerID>
    </{rq_name}>
  </QBXMLMsgsRq>
</QBXML>"""

    cont = f"""\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <{rq_name} requestID="1" iterator="Continue" iteratorID="{{iterator_id}}">
      <MaxReturned>100</MaxReturned>
    </{rq_name}>
  </QBXMLMsgsRq>
</QBXML>"""

    stop = f"""\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <{rq_name} requestID="1" iterator="Stop" iteratorID="{{iterator_id}}">
    </{rq_name}>
  </QBXMLMsgsRq>
</QBXML>"""

    return start, cont, stop


INVOICE_TEMPLATES    = _templates("InvoiceQueryRq")
CREDITMEMO_TEMPLATES = _templates("CreditMemoQueryRq")

CUSTOMER_QUERY = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <CustomerQueryRq requestID="1" iterator="Start">
      <MaxReturned>200</MaxReturned>
      <ActiveStatus>All</ActiveStatus>
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


def build_date_filter(from_date, to_date):
    parts = []
    if from_date:
        parts.append(f"<TxnDateRangeFilter><FromTxnDate>{from_date}</FromTxnDate>")
        if to_date:
            parts.append(f"<ToTxnDate>{to_date}</ToTxnDate>")
        parts.append("</TxnDateRangeFilter>")
    elif to_date:
        parts.append(f"<TxnDateRangeFilter><ToTxnDate>{to_date}</ToTxnDate></TxnDateRangeFilter>")
    return "".join(parts)


def text(node, tag):
    child = node.find(tag)
    return child.text.strip() if child is not None and child.text else ""


def extract_custom_field(data_ext_list, field_name):
    for ext in data_ext_list:
        if text(ext, "DataExtName") == field_name:
            return text(ext, "DataExtValue")
    return ""


def format_date(iso_date):
    """Convert YYYY-MM-DD to MM/DD/YYYY."""
    if iso_date and len(iso_date) == 10:
        y, m, d = iso_date.split("-")
        return f"{m}/{d}/{y}"
    return iso_date


def build_customer_map(backend):
    """Return a dict mapping CustomerRef ListID -> AccountNumber for all customers."""
    import xml.etree.ElementTree as ET

    customer_map = {}

    def _parse(xml_str):
        root = ET.fromstring(xml_str)
        rs = root.find(".//CustomerQueryRs")
        if rs is None:
            raise RuntimeError("No CustomerQueryRs in response")
        status_code = rs.attrib.get("statusCode", "0")
        if status_code not in ("0", "1"):
            raise RuntimeError(f"QB error {status_code}: {rs.attrib.get('statusMessage')}")
        for cust in rs.findall("CustomerRet"):
            list_id = text(cust, "ListID")
            account_number = text(cust, "AccountNumber")
            if list_id:
                customer_map[list_id] = account_number
        return rs.attrib.get("iteratorID", ""), int(rs.attrib.get("iteratorRemainingCount", "0"))

    iterator_id, remaining = _parse(backend.request(CUSTOMER_QUERY))
    print(f"  [CustomerRet] Fetched {len(customer_map)} customers, {remaining} remaining...")

    while remaining > 0 and iterator_id:
        iterator_id, remaining = _parse(
            backend.request(CUSTOMER_QUERY_CONT.format(iterator_id=iterator_id))
        )
        print(f"  [CustomerRet] Fetched {len(customer_map)} customers total, {remaining} remaining...")

    if iterator_id:
        backend.request(CUSTOMER_QUERY_STOP.format(iterator_id=iterator_id))

    print(f"  [CustomerRet] Customer map built: {len(customer_map)} entries")
    return customer_map


def parse_response(xml_str, rs_tag, ret_tag, ref_prefix, customer_map):
    import xml.etree.ElementTree as ET

    root = ET.fromstring(xml_str)
    rs = root.find(f".//{rs_tag}")
    if rs is None:
        raise RuntimeError(f"No {rs_tag} in response")

    status_code = rs.attrib.get("statusCode", "0")
    if status_code not in ("0", "1"):
        raise RuntimeError(f"QB error {status_code}: {rs.attrib.get('statusMessage')}")

    iterator_id = rs.attrib.get("iteratorID", "")
    remaining = int(rs.attrib.get("iteratorRemainingCount", "0"))

    rows = []
    for txn in rs.findall(ret_tag):
        list_id = text(txn, "CustomerRef/ListID")
        header_exts = txn.findall("DataExtRet")
        rows.append({
            "customer_id":      customer_map.get(list_id, list_id),
            "transaction_date": format_date(text(txn, "TxnDate")),
            "doc_num_h":        ref_prefix + text(txn, "RefNumber"),
            "amount":           text(txn, "SubTotal") or text(txn, "TotalAmount"),
            "stage":            extract_custom_field(header_exts, STAGE_FIELD_NAME),
            "transaction_id":   text(txn, "TxnNumber"),
            "open_balance":     text(txn, "BalanceRemaining"),
            "po_num":           text(txn, "PONumber"),
        })

    return rows, iterator_id, remaining


# ── Backends ───────────────────────────────────────────────────────────────────

class LocalBackend:
    def __init__(self):
        import win32com.client
        self.rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        self.rp.OpenConnection2("", "QB Invoice Header Extract", 1)
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


# ── Query runner ───────────────────────────────────────────────────────────────

def fetch_all(backend, templates, rs_tag, ret_tag, ref_prefix, date_filter, customer_map):
    start_tpl, cont_tpl, stop_tpl = templates

    all_rows = []
    response = backend.request(start_tpl.format(date_filter=date_filter))
    rows, iterator_id, remaining = parse_response(response, rs_tag, ret_tag, ref_prefix, customer_map)
    all_rows.extend(rows)
    print(f"  [{ret_tag}] Fetched {len(rows)} records, {remaining} remaining...")

    while remaining > 0 and iterator_id:
        response = backend.request(cont_tpl.format(iterator_id=iterator_id))
        rows, iterator_id, remaining = parse_response(response, rs_tag, ret_tag, ref_prefix, customer_map)
        all_rows.extend(rows)
        print(f"  [{ret_tag}] Fetched {len(rows)} records, {remaining} remaining...")

    if iterator_id:
        backend.request(stop_tpl.format(iterator_id=iterator_id))

    return all_rows


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Extract QB Desktop invoice/credit memo headers to CSV")
    parser.add_argument("--output", default="qbInvoiceHeader.csv", help="Output CSV filename")
    parser.add_argument("--output-dir", default=None,
                        help="Directory to write the output file (local path or network share)")
    parser.add_argument("--from-date", help="Start date YYYY-MM-DD (optional)")
    parser.add_argument("--to-date", help="End date YYYY-MM-DD (optional)")
    parser.add_argument("--qb-host", default=None,
                        help="IP/hostname of QB machine running qb_server.py (omit to run locally)")
    parser.add_argument("--qb-port", type=int, default=5000,
                        help="Port qb_server.py is listening on (default 5000)")
    args = parser.parse_args()

    if args.qb_host:
        backend = RemoteBackend(args.qb_host, args.qb_port)
        print(f"Remote mode: connecting to {args.qb_host}:{args.qb_port}")
    else:
        backend = LocalBackend()
        print("Local mode: connecting to QuickBooks on this machine")

    try:
        print("Building customer map...")
        customer_map = build_customer_map(backend)
        date_filter = build_date_filter(args.from_date, args.to_date)

        invoices = fetch_all(
            backend, INVOICE_TEMPLATES,
            "InvoiceQueryRs", "InvoiceRet", "Invoice - ", date_filter, customer_map,
        )
        credit_memos = fetch_all(
            backend, CREDITMEMO_TEMPLATES,
            "CreditMemoQueryRs", "CreditMemoRet", "Credit Memo - ", date_filter, customer_map,
        )
    finally:
        backend.close()

    all_rows = invoices + credit_memos
    all_rows.sort(key=lambda r: r["transaction_date"])

    output_path = os.path.join(args.output_dir, args.output) if args.output_dir else args.output

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"Done. {len(all_rows)} record(s) written to {output_path} "
          f"({len(invoices)} invoice, {len(credit_memos)} credit memo)")


if __name__ == "__main__":
    main()
