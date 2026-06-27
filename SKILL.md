---
name: snkr-prein
description: Generate warehouse pre-in spreadsheets from SNKRDUNK shipped orders. For each FedEx shipment from snkrdunk.com, produces one PreIn_<tracking>.xlsx (cloned from a user-supplied template) with one row per PSA-graded card. Use when a user mentions HORSTACC, FTL, SNKR warehouse intake, PreIn sheets, or generating xlsx for shipped SNKR orders.
---

# SNKR Pre-In Sheet Generator

For every shipped order on SNKRDUNK, the destination warehouse needs a pre-in spreadsheet keyed by the FedEx tracking number. This skill harvests tracking numbers + PSA cert lists from SNKRDUNK's order chat API and writes one xlsx per shipment.

## When to invoke

Use this skill when the user asks to:
- Generate PreIn / pre-in / warehouse intake sheets for SNKRDUNK shipments
- Produce xlsx files for new FedEx shipments from SNKR
- Sync SNKR shipping data to warehouse format (HORSTACC, FTL, or similar 3PL)

## Workflow

1. **Check config.** Look for `prein.config.json` in the working directory. If missing, copy `prein.config.example.json` and ask the user to fill required fields (`template`, `out_dir`, `warehouse`, `customer`, `collection`).
2. **Check auth.** Need `SNKRDUNK_ENSID` + `SNKRDUNK_SESSION` env vars. If missing, ask the user to paste them (from snkrdunk.com → DevTools → Application → Cookies → `ENSID` and `session` values).
3. **Harvest** with `node scripts/build-prein-shipments.mjs` — writes `prein-shipments.json` to cwd. Filters tracking by `tracking_regex` from config (default: all).
4. **Generate** with `python3 scripts/build-prein-xlsx.py` — writes one xlsx per shipment to `out_dir`. **Skips files that already exist** so re-runs only produce new shipments.

## Configuration (`prein.config.json`)

```json
{
  "template": "~/Downloads/PreIn_template.xlsx",
  "template_sheet": "Inbound - FTL",
  "out_dir": "~/Downloads",
  "out_prefix": "PreIn_",
  "out_suffix": "_updated",
  "tracking_regex": "",
  "customer": "FTL",
  "warehouse": "HORSTACC",
  "collection": "Pokemon",
  "grader": "PSA",
  "orders_csv_glob": "./orders-cart-*.csv",
  "psa_verify_glob": "./psa-verify-*.json",
  "columns": {
    "customer": 2,
    "inbound_date": 3,
    "tracking": 4,
    "warehouse": 6,
    "item_no": 7,
    "collection": 11,
    "card_name": 12,
    "grader": 13,
    "grade": 14,
    "pack_size": 15,
    "qty_ctns": 16,
    "qty_pcs": 17
  }
}
```

`columns` is a 1-indexed map from logical field → spreadsheet column number. Adjust to match your template. Set `tracking_regex` to e.g. `"^SF\\d+$"` to only process SF Express trackings.

## Cert backfill

SNKR's order API doesn't return PSA cert numbers directly. The skill backfills certs by matching each shipped order's `transaction_id` against local CSV files matching `orders_csv_glob`. Expected CSV columns: `transaction_id`, `psa_cert`, `grade`, `name`.

Optional: `psa_verify_glob` JSON files (output of a PSA cert verifier) supply canonical card descriptions and exact grade numbers, overriding the CSV values when present.

**Bundle orders:** if a SNKR transaction has qty>1, split it into one CSV row per physical card with its own `psa_cert`. Collapsed bundles will be skipped silently.

## Verify coverage

After step 1, confirm a cart's transactions are accounted for in the harvested shipments:

```bash
python3 scripts/check-cart-coverage.py <cart_id> [<cart_id>...]
```

Reports which trackings cover the cart and flags any tx not yet shipped.

## Requirements

- Node 18+ (built-in `fetch`)
- Python 3.10+ with `openpyxl` (`pip install openpyxl`)
- An xlsx template file (the skill clones it; doesn't generate styling from scratch)

## What gets populated per row

| Field | Source |
|---|---|
| Customer | config.customer |
| Inbound Date | tracking message sentAt |
| Tracking / Job No. | tracking number from `tracking number is [...]` |
| Warehouse | config.warehouse |
| Item No. | `PSA-<cert>` |
| Collection | config.collection |
| Card Name | PSA canonical description (psa-verify) or SNKR name with `[set]`/`(pack)` stripped |
| Grader | config.grader |
| Grade | exact grade from psa-verify, or parsed from CSV `grade` |
| Pack Size / Qty | 1 / 1 / 1 |

## Notes

- The two scripts are independent — re-run step 2 alone to retry generation without re-hitting SNKR's API.
- SNKR API endpoints used: `/en/v1/account/orders`, `/en/v1/account/orders/cart/{cartId}`, `/en/v1/account/orders/{orderId}/contacts`. These are unofficial; behavior may change without notice.
- Cookies expire — refresh them when step 1 returns 401/403.
