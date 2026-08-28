# Cow-Calf Rancher Lookup

A single Excel tool that, for a given county, brings together four cattle-ranching
datasets for 2025:

- **Auction cattle prices** — per-state head-weighted averages from USDA Market
  News weekly auction summaries (feeder steers/heifers/bulls, bred cows/heifers,
  cow-calf & heifer pairs, stock cows, replacement bulls), in `$/cwt` or `$/head`.
- **Hay price** — average `$/ton` by state.
- **Calf margin** — `$/cow` by the county's USDA ERS farm resource region.
- **Fencing estimate** — build cost by fence type and length (acres, miles, or feet).

## The combined tool: `Rancher_Lookup_Combined.xlsx`

Open the **Rancher Lookup** tab, pick a **state** and then a **county**, and the
sheet fills in the auction cattle averages and hay price (by state) and the calf
margin (by the county's ERS region). The **fencing estimator** is independent:
choose a fence type and an amount in acres, miles, or feet to get a build cost.

Behind the scenes, one data tab holds each dataset (`Cattle_by_State`,
`Hay_by_State`, `CalfMargin_by_Region`, `Fencing_Rates`) and a `Counties` tab
maps every U.S. county to its state and ERS region, so a single county choice
drives every metric. All values are live `XLOOKUP`/`FILTER` formulas.

**Coverage:** auction cattle prices cover 25 states and hay covers 27; a county
outside those shows `n/a` for that metric, while calf margin and fencing cover
every county. For **acres**, fence length assumes a square paddock
(perimeter = 4 × √(acres × 43,560 sq ft)); **feet** and **miles** are exact.

## Standalone workbooks

Kept alongside the combined tool for reference — each is the original single-topic
lookup those numbers came from:

- `USDA_Cattle_2025_State_Averages.xlsx` — the cattle averages on their own, with
  a state/category lookup and a per-number audit tab.
- `Calf Margin by County.xlsx` — the county → ERS region → margin lookup.
- `HayPriceLookup.xlsx` — hay price by state.
- `Livestock Fencing Cost Calculator.xlsx` — fencing build cost by type and length.

## Rebuilding from source

Requires Python and `pip install -r requirements.txt`.

The combined workbook and the cattle averages rebuild from the committed `data/`
files with no network access:

```bash
python scripts/2_compute.py          # cattle averages     -> data/averages_long.csv
python scripts/3_build_xlsx.py       # cattle workbook     -> USDA_Cattle_2025_State_Averages.xlsx
python scripts/4_build_combined.py   # combined tool       -> Rancher_Lookup_Combined.xlsx
```

To re-pull the cattle data from USDA (e.g. for a new year), set a USDA Market
News (MARS) API key and run the first two steps:

```bash
export AMS_API_KEY="your-key"
python scripts/0_discover_reports.py   # report IDs + category inventory
python scripts/1_extract.py            # pull + filter weekly rows -> data/cattle_2025.csv.gz
```

## Data files

- `data/counties.csv` — county → state + FIPS + ERS region (3,144 counties).
- `data/combine_data.json` — hay price by state, fencing `$/foot` by type, calf
  margin by ERS region.
- `data/averages_long.csv` — computed per-state per-category cattle averages.
- `data/cattle_2025.csv.gz` — filtered 2025 auction observation rows (the extract).
- `data/report_inventory.json` — discovered USDA reports and category combinations.
