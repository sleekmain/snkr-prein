#!/usr/bin/env python3
"""For a given cart ID, list which FedEx shipments cover its transactions.

Usage:  python3 scripts/check-cart-coverage.py <cart_id> [<cart_id>...]
Needs:  prein-shipments.json + orders-cart-<cart>.csv in cwd.
"""
import csv
import json
import sys
from pathlib import Path

if len(sys.argv) < 2:
    sys.exit("usage: check-cart-coverage.py <cart_id> [<cart_id>...]")

ships = json.loads(Path("prein-shipments.json").read_text())

for cart in sys.argv[1:]:
    csv_path = Path(f"orders-cart-{cart}.csv")
    if not csv_path.exists():
        print(f"cart {cart}: no CSV ({csv_path})")
        continue
    txs = set()
    with open(csv_path) as f:
        for r in csv.DictReader(f):
            tx = (r.get("transaction_id") or "").strip()
            if tx:
                txs.add(tx)
    print(f"cart {cart}: {len(txs)} tx total")

    hits = {}
    for tracking, ship in ships.items():
        for row in ship["rows"]:
            if row.get("tx") in txs:
                hits.setdefault(tracking, []).append(row["tx"])
    if not hits:
        print("  (no shipments cover this cart yet)")
        continue
    covered = sum(len(v) for v in hits.values())
    for t, items in sorted(hits.items()):
        print(f"  → {t}: {len(items)} items")
    if covered < len(txs):
        print(f"  ⚠ {len(txs) - covered} tx not yet in any shipment (unshipped or missing chat tracking)")
