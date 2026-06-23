# snkr-prein

A [Claude Code](https://docs.claude.com/en/docs/claude-code) skill that generates warehouse pre-in spreadsheets from your SNKRDUNK shipped orders.

For every FedEx shipment from snkrdunk.com, it produces one `PreIn_<tracking>.xlsx` (cloned from a template you supply) with one row per PSA-graded card. Designed for buyers who consolidate SNKR purchases through a 3PL (e.g. HORSTACC) that needs a pre-inbound sheet for each package.

## Install

```bash
mkdir -p ~/.claude/skills
git clone https://github.com/<your-user>/snkr-prein.git ~/.claude/skills/snkr-prein
cd ~/.claude/skills/snkr-prein
pip install openpyxl
```

Claude Code auto-discovers the skill on next launch. Invoke by asking Claude to "generate prein sheets" or by referencing SNKR / HORSTACC / warehouse intake.

## Setup

1. Copy `prein.config.example.json` → `prein.config.json` in the directory where you'll run the scripts. Edit:
   - `template` — path to your warehouse's pre-in xlsx template
   - `template_sheet` — sheet name inside the template
   - `customer`, `warehouse`, `collection`, `grader` — your warehouse defaults
   - `columns` — 1-indexed column numbers for each field in your template
   - `tracking_regex` — optional filter (e.g. `"^SF\\d+$"` for SF Express only)

2. Get your SNKRDUNK cookies: snkrdunk.com → DevTools → Application → Cookies → copy `ENSID` and `session` values.

3. Run:
   ```bash
   SNKRDUNK_ENSID=... SNKRDUNK_SESSION=... node scripts/build-prein-shipments.mjs
   python3 scripts/build-prein-xlsx.py
   ```

   Step 1 writes `prein-shipments.json` (tracking → orders). Step 2 writes one xlsx per shipment to `out_dir`. **Existing files are never overwritten**, so re-runs only produce new shipments.

## How cert backfill works

SNKRDUNK's order API doesn't return PSA cert numbers directly. The skill backfills certs by matching each shipped order's `transaction_id` against local CSV files matching `orders_csv_glob`. Expected CSV columns:

| transaction_id | psa_cert | grade | name |
|---|---|---|---|
| 1001033-JP | 148337021 | PSA 9 | Swinub AR [SV9 106/100]... |

You're responsible for building these CSVs (typically by scraping your cart pages — that part isn't included here since it depends on OCR/manual entry).

Optional `psa_verify_glob` JSON files (output of a PSA cert verifier) supply canonical card descriptions and exact grade numbers, overriding the CSV values when present.

## Requirements

- Node 18+ (built-in `fetch`)
- Python 3.10+ with `openpyxl`

## Disclaimers

SNKRDUNK API endpoints used here are unofficial. Behavior may change without notice. Cookies expire — refresh them when step 1 returns 401/403. Use at your own risk and within SNKRDUNK's Terms of Service.

## License

MIT
