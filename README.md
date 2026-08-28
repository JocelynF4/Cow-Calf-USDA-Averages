# Cow-Calf USDA Averages

Per-state **2025 annual average auction prices** for cattle, built from USDA
Market News weekly auction summary reports (MARS API v1.2).

The deliverable is **`USDA_Cattle_2025_State_Averages.xlsx`**:

- **Data** — a 25-state × 15-category matrix of head-weighted annual averages,
  with per-state week counts and a partial-year coverage flag.
- **Lookup** — pick a state and a category from dropdowns to see that state's
  2025 average, its unit, and the weeks/head of data behind it.
- **Source Detail** — every state × category number with its total head,
  observation count, and week count, for auditing.

## Method

- **Source:** each state's `... Weekly Cattle/Livestock Auction Summary` report,
  pulled for the full range `01/01/2025–12/31/2025` (every published weekly issue).
- **Categories:** Feeder Cattle and Replacement Cattle only — no slaughter, no
  dairy. Steers, Heifers, Bulls, Cows (feeder); Stock Cows, Bred Cows, Bred
  Heifers, Cow-Calf Pairs, Heifer Pairs, breeding Bulls (replacement).
- **Averaging:** head-count-weighted across every matching row in the year:
  `Σ(avg_price × head_count) / Σ(head_count)`.
- **Overall:** head-weighted blend of feeder Steers + Heifers + Bulls ($/cwt).
- **Units:** categories priced per hundredweight are `$/cwt`; bred stock and
  pairs are priced `$/head`. Categories reported both ways across states appear
  as two columns; the two units are never mixed.
- **Coverage:** West Virginia published no 2025 data (25 states in the output).
  Some states report fewer than 52 weeks; those are flagged as partial-year.

## Reproduce

Requires a USDA Market News (MARS) API key in the environment:

```bash
export AMS_API_KEY="your-key"
pip install -r requirements.txt

python scripts/0_discover_reports.py   # report IDs + category inventory -> data/report_inventory.json
python scripts/1_extract.py            # pull + filter 2025 rows        -> data/cattle_2025.csv.gz
python scripts/2_compute.py            # head-weighted averages         -> data/averages_long.csv
python scripts/3_build_xlsx.py         # workbook                        -> USDA_Cattle_2025_State_Averages.xlsx
```

Steps 2 and 3 run offline from the committed `data/` files; only steps 0 and 1
call the API.

## Data files

- `data/cattle_2025.csv.gz` — filtered 2025 observation rows (the extract).
- `data/averages_long.csv` — computed per-state per-category averages.
- `data/report_inventory.json` — discovered state reports and the full list of
  category combinations found in the 2025 data.
