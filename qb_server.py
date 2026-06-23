"""
QuickBooks Desktop - QBXML Relay Server
Runs on the QB machine (192.168.0.81) and exposes a local HTTP API so that
qb_invoice_extract.py can be run from any machine on the LAN.

Requirements (QB machine only):
  pip install pywin32 flask

Usage:
  python qb_server.py [--port 5000] [--host 0.0.0.0]

Security note: bind to the LAN interface only and firewall the port from WAN.
"""

import argparse
import win32com.client
from flask import Flask, request, jsonify

app = Flask(__name__)

_rp = None
_ticket = None


def get_session():
    global _rp, _ticket
    if _rp is None:
        _rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        _rp.OpenConnection2("", "QB Invoice Extract", 1)
        _ticket = _rp.BeginSession("", 2)
    return _rp, _ticket


@app.route("/qbxml", methods=["POST"])
def qbxml():
    xml = request.data.decode("utf-8")
    if not xml.strip():
        return jsonify({"error": "empty request body"}), 400
    try:
        rp, ticket = get_session()
        response = rp.ProcessRequest(ticket, xml)
        return response, 200, {"Content-Type": "text/xml; charset=utf-8"}
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


def main():
    parser = argparse.ArgumentParser(description="QB QBXML relay server")
    parser.add_argument("--host", default="0.0.0.0", help="Bind address (default 0.0.0.0)")
    parser.add_argument("--port", type=int, default=5000, help="Port (default 5000)")
    args = parser.parse_args()

    # Open QB session at startup so the user sees the access prompt immediately.
    print("Connecting to QuickBooks...")
    get_session()
    print(f"QB session open. Listening on {args.host}:{args.port}")
    app.run(host=args.host, port=args.port)


if __name__ == "__main__":
    main()
