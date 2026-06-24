"""
QuickBooks Desktop - Products / Items Catalog Extract
Outputs CSV matching the qbProducts format.

Local mode (run directly on the QB machine):
  python qb_products_extract.py [--output qbProducts.csv]
  Requirements: Windows + QuickBooks Desktop open, pip install pywin32

Remote mode (run from any machine on the LAN):
  python qb_products_extract.py --qb-host 192.168.0.81 [--qb-port 5000] [options]
  Requirements: pip install requests
  The QB machine must be running qb_server.py.
"""

import csv
import argparse
import os

COLUMNS = [
    "id",
    "name",
    "sales_price_amt",
    "on_hand_qnty",
    "on_po_order_qnty",
    "on_so_order_qnty",
    "unit_cost_amt",
    "description",
    "is_active",
    "premier_name",
    "premier_price",
    "dist_name",
    "dist_price",
    "standard_name",
]

# Price level names as they appear in QuickBooks
PRICE_LEVEL_PREMIER  = "Premier"
PRICE_LEVEL_DIST     = "Distributor"
PRICE_LEVEL_STANDARD = "Standard Price Book"

QBXML_START = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <ItemQueryRq requestID="1" iterator="Start">
      <MaxReturned>200</MaxReturned>
      <ActiveStatus>All</ActiveStatus>
      <OwnerID>0</OwnerID>
    </ItemQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

QBXML_CONTINUE = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <ItemQueryRq requestID="1" iterator="Continue" iteratorID="{iterator_id}">
      <MaxReturned>200</MaxReturned>
    </ItemQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

QBXML_STOP = """\
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <ItemQueryRq requestID="1" iterator="Stop" iteratorID="{iterator_id}">
    </ItemQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

# All item return element types QB may include in an ItemQueryRs
ITEM_RET_TAGS = [
    "ItemServiceRet",
    "ItemNonInventoryRet",
    "ItemInventoryRet",
    "ItemInventoryAssemblyRet",
    "ItemFixedAssetRet",
    "ItemOtherChargeRet",
    "ItemDiscountRet",
    "ItemPaymentRet",
    "ItemSalesTaxRet",
    "ItemSalesTaxGroupRet",
    "ItemGroupRet",
    "ItemSubtotalRet",
]


def text(node, tag):
    child = node.find(tag)
    return child.text.strip() if child is not None and child.text else ""


def price_level_value(item_node, level_name):
    """Return the per-item price for a named price level, or '' if not set."""
    for pl in item_node.findall("PriceLevelPerItemRet"):
        if text(pl, "PriceLevelRef/FullName") == level_name:
            return text(pl, "Price")
    return ""


def parse_item(node, row_id):
    """Extract a product row from any item *Ret element."""
    # Sales price: inventory / non-inventory split into SalesOrPurchase vs SalesAndPurchase
    sales_price = (
        text(node, "SalesPrice")
        or text(node, "SalesOrPurchaseInfo/Price")
        or text(node, "SalesAndPurchaseInfo/SalesPrice")
        or "0.00"
    )

    # Unit cost
    unit_cost = (
        text(node, "PurchaseCost")
        or text(node, "SalesOrPurchaseInfo/Price")   # service-only items reuse same field
        or text(node, "SalesAndPurchaseInfo/PurchaseCost")
        or "0.00"
    )
    # For service/other-charge items that have only one price field, cost should be 0
    # if the tag we found was the sales price field.
    if node.tag in ("ItemServiceRet", "ItemOtherChargeRet", "ItemSalesTaxRet",
                    "ItemDiscountRet", "ItemPaymentRet", "ItemSubtotalRet",
                    "ItemSalesTaxGroupRet", "ItemGroupRet"):
        unit_cost = "0.00"

    description = (
        text(node, "SalesDesc")
        or text(node, "SalesOrPurchaseInfo/Desc")
        or text(node, "SalesAndPurchaseInfo/SalesDesc")
        or text(node, "Desc")
        or ""
    )

    return {
        "id":               row_id,
        "name":             text(node, "FullName") or text(node, "Name"),
        "sales_price_amt":  sales_price,
        "on_hand_qnty":     text(node, "QuantityOnHand"),
        "on_po_order_qnty": text(node, "QuantityOnPurchaseOrder"),
        "on_so_order_qnty": text(node, "QuantityOnSalesOrder"),
        "unit_cost_amt":    unit_cost,
        "description":      description,
        "is_active":        text(node, "IsActive") or "True",
        "premier_name":     PRICE_LEVEL_PREMIER,
        "premier_price":    price_level_value(node, PRICE_LEVEL_PREMIER),
        "dist_name":        PRICE_LEVEL_DIST,
        "dist_price":       price_level_value(node, PRICE_LEVEL_DIST),
        "standard_name":    PRICE_LEVEL_STANDARD,
    }


def parse_response(xml_str):
    import xml.etree.ElementTree as ET

    root = ET.fromstring(xml_str)
    rs = root.find(".//ItemQueryRs")
    if rs is None:
        raise RuntimeError("No ItemQueryRs in response")

    status_code = rs.attrib.get("statusCode", "0")
    if status_code not in ("0", "1"):
        raise RuntimeError(f"QB error {status_code}: {rs.attrib.get('statusMessage')}")

    iterator_id = rs.attrib.get("iteratorID", "")
    remaining = int(rs.attrib.get("iteratorRemainingCount", "0"))

    items = []
    for tag in ITEM_RET_TAGS:
        items.extend(rs.findall(tag))

    return items, iterator_id, remaining


# ── Backends ───────────────────────────────────────────────────────────────────

class LocalBackend:
    def __init__(self):
        import win32com.client
        self.rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        self.rp.OpenConnection2("", "QB Products Extract", 1)
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


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Extract QB Desktop item/product catalog to CSV")
    parser.add_argument("--output", default="qbProducts.csv", help="Output CSV filename")
    parser.add_argument("--output-dir", default=None,
                        help="Directory to write the output file (local path or network share)")
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
        all_items = []
        response = backend.request(QBXML_START)
        items, iterator_id, remaining = parse_response(response)
        all_items.extend(items)
        print(f"  Fetched {len(items)} items, {remaining} remaining...")

        while remaining > 0 and iterator_id:
            response = backend.request(QBXML_CONTINUE.format(iterator_id=iterator_id))
            items, iterator_id, remaining = parse_response(response)
            all_items.extend(items)
            print(f"  Fetched {len(items)} items, {remaining} remaining...")

        if iterator_id:
            backend.request(QBXML_STOP.format(iterator_id=iterator_id))

    finally:
        backend.close()

    rows = [parse_item(node, idx + 1) for idx, node in enumerate(all_items)]

    output_path = os.path.join(args.output_dir, args.output) if args.output_dir else args.output

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Done. {len(rows)} item(s) written to {output_path}")


if __name__ == "__main__":
    main()
