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
drives every metric. All values are live `INDEX`/`MATCH` formulas.

**Coverage:** auction cattle prices cover 25 states and hay covers 27; a county
outside those shows `n/a` for that metric, while calf margin and fencing cover
every county. For **acres**, fence length assumes a square paddock
(perimeter = 4 × √(acres × 43,560 sq ft)); **feet** and **miles** are exact.

## How each source is averaged

Each of the four datasets is aggregated differently, because each is published
at a different level of detail. Here is exactly what happens to each one.

### 1. Auction cattle prices — head-count-weighted average, by state

- **Source:** every 2025 weekly auction summary report USDA published for each
  state (25 states report; West Virginia published none in 2025).
- **What the raw data looks like:** each weekly report splits a class (e.g.
  feeder steers) into many rows by weight break, frame, muscle grade, etc. —
  and every row carries its own average price *and* a head count.
- **The average:** for each state and category, all of that state's matching
  rows across all 2025 weeks are combined into one **head-count-weighted mean**:

  ```
  average = Σ (row average price × row head count) ÷ Σ (row head count)
  ```

  So a week (or a weight class) that sold 800 head counts 800× as much as one
  that sold 1 head. This reflects the price actual volume traded at, rather than
  averaging small and large sales equally.
- **"Overall Feeder"** uses the same formula, pooling feeder **Steers +
  Heifers + Bulls** together.
- **Units:** classes priced by weight are `$/cwt`; bred stock and pairs are
  `$/head`. A few categories are reported both ways across different states —
  those are kept as **two separate columns** and never blended together.
- **Excluded from the average:** rows with no price or zero head count, and all
  slaughter and dairy classes.

### 2. Hay price — simple monthly average, by state

- **Source:** USDA NASS hay price survey — **five monthly prices** (June through
  October 2025) for each of **27 states**.
- **The average:** for each state, the plain arithmetic mean of those five
  monthly `$/ton` prices (each month weighted equally). This is **not**
  volume-weighted — it is a straight average of the monthly survey figures.

### 3. Calf margin — a regional value, not re-averaged

- **Source:** the calf-margin workbook's 2025 margin (`$/cow`) for each of
  USDA's **9 ERS farm resource regions**.
- **What the tool does:** **no averaging happens in the combined tool.** Each
  county is simply assigned its ERS region's single margin value, using the
  county → region crosswalk. Every county in the same region shows the same
  number. (How that regional margin was originally derived lives in the
  standalone `Calf Margin by County.xlsx` tool.)

### 4. Fencing — a per-foot rate × length, not averaged

- **Source:** the fencing calculator's **cost per foot** for each of **5 fence
  types**, taken as each type's total itemized build cost for a standard fence
  segment divided by that segment's length.
- **What the tool does:** **no averaging across geography** — fencing does not
  vary by state or county. The estimate is simply:

  ```
  total cost = fence length (feet) × cost per foot for the chosen type
  ```

- **Length conversion:** feet are used as entered; **miles × 5,280**; **acres**
  are converted to a fence length by assuming a **square paddock**
  (perimeter = 4 × √(acres × 43,560 sq ft)) — an estimate, since an acreage
  alone does not fix a field's shape. Feet and miles are exact.

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

## Source data (original uploads)

The raw inputs the tools were built from, kept for reference:

- `Hay prices state June-Sept.csv` — USDA NASS hay price survey rows.
- `livestock_fencing_costs.csv` — itemized fencing costs by fence type.
- `uscounties.xlsx` — county reference (name, FIPS, state).
- `ers_farm_resource_regions_county_crosswalk.xls` — county → USDA ERS region crosswalk.
- `CowCalfCostReturn (1).xlsx` — USDA cow-calf cost-and-return survey.
- `hay_prices.py` — script behind the hay lookup.

## Data files

- `data/counties.csv` — county → state + FIPS + ERS region (3,144 counties).
- `data/combine_data.json` — hay price by state, fencing `$/foot` by type, calf
  margin by ERS region.
- `data/averages_long.csv` — computed per-state per-category cattle averages.
- `data/cattle_2025.csv.gz` — filtered 2025 auction observation rows (the extract).
- `data/report_inventory.json` — discovered USDA reports and category combinations.
