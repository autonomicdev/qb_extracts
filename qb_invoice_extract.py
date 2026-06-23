"""
QuickBooks Desktop - Invoice Detail Extract with Serial Numbers
Outputs CSV matching the qbInvoiceDetailSN format.

Requirements:
  - Windows with QuickBooks Desktop open and company file loaded
  - pip install pywin32
  - Run as a user with QB access (QB will prompt to allow access on first run)

Usage:
  python qb_invoice_extract.py [--output invoices.csv] [--from-date YYYY-MM-DD] [--to-date YYYY-MM-DD]
"""

import csv
import sys
import argparse
import win32com.client

COLUMNS = [
    "InvoiceLineTxnLineID",
    "TxnDate",
    "TimeModified",
    "TxnNumber",
    "RefNumber",
    "InvoiceLineSeqNo",
    "InvoiceLineDesc",
    "InvoiceLineQuantity",
    "InvoiceLineRate",
    "InvoiceLineAmount",
    "InvoiceLineSerialNumber",
    "InvoiceLineTaxAmount",
    "InvoiceLineItemRefFullName",
    "Commissionable",
    "PONumber",
]

QBXML_TEMPLATE = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <InvoiceQueryRq requestID="1" iterator="Start">
      <MaxReturned>100</MaxReturned>
      {date_filter}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </InvoiceQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

QBXML_CONTINUE = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <InvoiceQueryRq requestID="1" iterator="Continue" iteratorID="{iterator_id}">
      <MaxReturned>100</MaxReturned>
    </InvoiceQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

QBXML_CLOSE = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <InvoiceQueryRq requestID="1" iterator="Stop" iteratorID="{iterator_id}">
    </InvoiceQueryRq>
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
    """Return text content of first matching child element, or empty string."""
    import xml.etree.ElementTree as ET
    child = node.find(tag)
    return child.text.strip() if child is not None and child.text else ""


def extract_custom_field(data_ext_list, field_name):
    """Get value of a DataExt custom field by name."""
    import xml.etree.ElementTree as ET
    for ext in data_ext_list:
        if text(ext, "DataExtName") == field_name:
            return text(ext, "DataExtValue")
    return ""


def parse_response(xml_str):
    """Parse QBXML InvoiceQueryRs and return (rows, iterator_id, remaining)."""
    import xml.etree.ElementTree as ET

    root = ET.fromstring(xml_str)
    rs = root.find(".//InvoiceQueryRs")
    if rs is None:
        raise RuntimeError("No InvoiceQueryRs in response")

    status_code = rs.attrib.get("statusCode", "0")
    if status_code not in ("0", "1"):  # 1 = no records, still ok
        raise RuntimeError(f"QB error {status_code}: {rs.attrib.get('statusMessage')}")

    iterator_id = rs.attrib.get("iteratorID", "")
    remaining = int(rs.attrib.get("iteratorRemainingCount", "0"))

    rows = []
    for inv in rs.findall("InvoiceRet"):
        txn_date = text(inv, "TxnDate")
        time_modified = text(inv, "TimeModified")
        txn_number = text(inv, "TxnNumber")
        ref_number = text(inv, "RefNumber")
        po_number = text(inv, "PONumber")

        # Header-level custom fields
        header_exts = inv.findall("DataExtRet")
        commissionable_hdr = extract_custom_field(header_exts, "Commissionable")

        for line in inv.findall("InvoiceLineRet"):
            line_exts = line.findall("DataExtRet")
            rows.append({
                "InvoiceLineTxnLineID":       text(line, "TxnLineID"),
                "TxnDate":                    txn_date,
                "TimeModified":               time_modified,
                "TxnNumber":                  txn_number,
                "RefNumber":                  ref_number,
                "InvoiceLineSeqNo":           text(line, "SeqNo"),
                "InvoiceLineDesc":            text(line, "Desc"),
                "InvoiceLineQuantity":        text(line, "Quantity"),
                "InvoiceLineRate":            text(line, "Rate"),
                "InvoiceLineAmount":          text(line, "Amount"),
                "InvoiceLineSerialNumber":    extract_custom_field(line_exts, "Serial Number"),
                "InvoiceLineTaxAmount":       text(line, "SalesTaxAmount"),
                "InvoiceLineItemRefFullName": text(line, "ItemRef/FullName"),
                "Commissionable":             extract_custom_field(line_exts, "Commissionable") or commissionable_hdr,
                "PONumber":                   po_number,
            })

    return rows, iterator_id, remaining


def run_query(rp, ticket, xml):
    response = rp.ProcessRequest(ticket, xml)
    return response


def main():
    parser = argparse.ArgumentParser(description="Extract QB Desktop invoice detail to CSV")
    parser.add_argument("--output", default="qbInvoiceDetailSN.csv", help="Output CSV file path")
    parser.add_argument("--from-date", help="Start date YYYY-MM-DD (optional)")
    parser.add_argument("--to-date", help="End date YYYY-MM-DD (optional)")
    args = parser.parse_args()

    rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    rp.OpenConnection2("", "QB Invoice Extract", 1)  # 1 = localQBD
    ticket = rp.BeginSession("", 2)  # 2 = doNotCare (uses open company file)

    try:
        date_filter = build_date_filter(args.from_date, args.to_date)
        first_xml = QBXML_TEMPLATE.format(date_filter=date_filter)

        all_rows = []
        response = run_query(rp, ticket, first_xml)
        rows, iterator_id, remaining = parse_response(response)
        all_rows.extend(rows)
        print(f"  Fetched {len(rows)} lines, {remaining} invoices remaining...")

        while remaining > 0 and iterator_id:
            response = run_query(rp, ticket, QBXML_CONTINUE.format(iterator_id=iterator_id))
            rows, iterator_id, remaining = parse_response(response)
            all_rows.extend(rows)
            print(f"  Fetched {len(rows)} lines, {remaining} invoices remaining...")

        if iterator_id:
            rp.ProcessRequest(ticket, QBXML_CLOSE.format(iterator_id=iterator_id))

    finally:
        rp.EndSession(ticket)
        rp.CloseConnection()

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"Done. {len(all_rows)} line(s) written to {args.output}")


if __name__ == "__main__":
    main()
