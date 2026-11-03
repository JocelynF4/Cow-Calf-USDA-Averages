# Cow-Calf Rancher Lookup

A single Excel tool that, for a given county, brings together five cattle-ranching
datasets:

- **Auction cattle prices** — per-state head-weighted averages from USDA Market
  News weekly auction summaries (feeder steers/heifers/bulls, bred cows/heifers,
  cow-calf & heifer pairs, stock cows, replacement bulls), in `$/cwt` or `$/head`.
- **Hay price** — average `$/ton` by state.
- **Calf margin** — `$/cow` by the county's USDA ERS farm resource region.
- **Fencing estimate** — build cost by fence type and length (acres, miles, or feet).
- **Tractor variable cost per hour** — fuel, lubrication and repairs (wear and tear)
  per hour for a two-wheel-drive (2WD) tractor, using the state's regional diesel price.

## The combined tool: `Rancher_Lookup_Combined.xlsx`

Open the **Rancher Lookup** tab, pick a **state** and then a **county**, and the
sheet fills in the auction cattle averages and hay price (by state) and the calf
margin (by the county's ERS region). The **fencing estimator** is independent:
choose a fence type and an amount in acres, miles, or feet to get a build cost.
The **tractor variable cost per hour** section uses the chosen state's regional diesel
price; fill in the tractor's horsepower, lubrication cost, list price and hours.

Behind the scenes, one data tab holds each dataset (`Cattle_by_State`,
`Hay_by_State`, `CalfMargin_by_Region`, `Fencing_Rates`) and a `Counties` tab
maps every U.S. county to its state and ERS region, so a single county choice
drives every metric. All values are live `INDEX`/`MATCH` formulas.

**Coverage:** auction cattle prices cover 25 states and hay covers 27; a county
outside those shows `n/a` for that metric, while calf margin and fencing cover
every county. For **acres**, fence length assumes a square paddock
(perimeter = 4 × √(acres × 43,560 sq ft)); **feet** and **miles** are exact.

## How each source is averaged

Each of the five datasets is aggregated differently, because each is published
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

### 5. Tractor variable cost per hour — regional diesel price + a standard repair model

**Diesel price — simple weekly average, by fuel region**

- **Source:** `Weekly_On-Highway_Diesel_Fuel_Prices_20261103.csv` — weekly retail
  on-highway diesel `$/gal` for the U.S. and each fuel region, Nov 2021 – 2 Nov 2026.
- **State → region:** every state is assigned to its Petroleum Administration for
  Defense District (PADD), and the matching price series is used:

  | PADD | Region | States | Price series used |
  |---|---|---|---|
  | 1A | New England | CT, ME, MA, NH, RI, VT | New England |
  | 1B | Central Atlantic | DE, DC, MD, NJ, NY, PA | Central Atlantic |
  | 1C | Lower Atlantic | FL, GA, NC, SC, VA, WV | Lower Atlantic |
  | 2 | Midwest | IL, IN, IA, KS, KY, MI, MN, MO, NE, ND, OH, OK, SD, TN, WI | Midwest |
  | 3 | Gulf Coast | AL, AR, LA, MS, NM, TX | Gulf Coast |
  | 4 | Rocky Mountain | CO, ID, MT, UT, WY | Rocky Mountains |
  | 5 | West Coast | CA | California |
  | 5 | West Coast | AK, AZ, HI, NV, OR, WA | West Coast exc California |
  | 6 | Caribbean territories | Puerto Rico, U.S. Virgin Islands | none published |
  | 7 | Pacific territories | Guam, American Samoa, Northern Mariana Islands | none published |

  PADD 1 uses its three sub-district series rather than the blended East Coast
  figure, and California has its own series, so both are used for accuracy.
- **The average:** a plain arithmetic mean of the weekly prices (each week weighted
  equally). Four bases are kept and chosen from a dropdown: **5-year average**
  (default; the 260 weeks from 15 Nov 2021 to 2 Nov 2026), **52-week average**
  (10 Nov 2025 – 2 Nov 2026), **latest week** (2 Nov 2026), and the **2025
  calendar-year average** (to line up with the other 2025 datasets). Prices rose
  sharply through 2026, so the latest week is well above the averages
  (U.S.: latest $6.29, 52-week $4.93, 5-year $4.31).
- **Caveats:** these are on-highway retail prices, which include road fuel taxes
  that off-road (dyed) farm diesel does not pay, so they overstate farm fuel cost.
  Enter your own price to override. Territories have no series and show `n/a`
  until a price is entered.

**Variable cost per hour — 2WD tractors only**

Only **variable costs** are measured: costs that rise with every hour the tractor
runs. Ownership costs (depreciation, interest, taxes, insurance and housing) and
operator labor are **not included**. The model covers **two-wheel-drive (2WD)
tractors only**.

```
fuel        = 0.044 gal per hp-hour × max PTO hp × diesel price
lubrication = entered directly in $/hr (oil, filters, grease)
repairs     = list price × RF1 × [((hours on + hours/yr)/1000)^RF2 − (hours on/1000)^RF2] ÷ hours/yr
total       = fuel + lubrication + repairs
```

- **Fuel use** (0.044 gal per horsepower-hour) is the average diesel use from Iowa
  State University's *Estimating Farm Machinery Costs* (Ag Decision Maker A3-29).
- **Lubrication** is its own input and is **not tied to the diesel price**. The
  $3.00/hr default is a placeholder; replace it with your own oil, filter and grease cost.
- **Repairs & maintenance** (reported as wear and tear) use the 2WD tractor repair
  factors from the American Society of Agricultural and Biological Engineers
  standard D497 (RF1 = 0.007, RF2 = 2.0). Repairs grow with accumulated hours, so an
  older, high-hour tractor costs more per hour to repair than a new one. The
  formula costs this year's hours on top of the hours already on the tractor.
- Factors live on the `Tractor_Factors` tab and can be edited. They were entered
  from memory of the published sources and have not yet been checked against them.
- Defaults (100 hp, $3.00/hr lubrication, $120,000 list price, 3,000 hours on,
  400 hours/yr) are placeholders; replace them with your own.

## Standalone workbooks

Kept alongside the combined tool for reference — each is the original single-topic
lookup those numbers came from:

- `USDA_Cattle_2025_State_Averages.xlsx` — the cattle averages on their own, with
  a state/category lookup and a per-number audit tab.
- `Calf Margin by County.xlsx` — the county → ERS region → margin lookup.
- `HayPriceLookup.xlsx` — hay price by state.
- `Livestock Fencing Cost Calculator.xlsx` — fencing build cost by type and length.
- `Tractor Cost per Hour.xlsx` — 2WD tractor variable cost per hour by state; also covers the
  territories (PADD 6 and 7) with a manual diesel price.

## Rebuilding from source

Requires Python and `pip install -r requirements.txt`.

The combined workbook and the cattle averages rebuild from the committed `data/`
files with no network access:

```bash
python scripts/2_compute.py            # cattle averages     -> data/averages_long.csv
python scripts/3_build_xlsx.py         # cattle workbook     -> USDA_Cattle_2025_State_Averages.xlsx
python scripts/5_diesel_by_region.py   # diesel averages     -> data/diesel_by_region.csv
python scripts/4_build_combined.py     # combined tool       -> Rancher_Lookup_Combined.xlsx
python scripts/6_build_tractor_cost.py # tractor workbook    -> Tractor Cost per Hour.xlsx
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
- `Weekly_On-Highway_Diesel_Fuel_Prices_20261103.csv` — weekly on-highway diesel prices by fuel region.
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
- `data/diesel_by_region.csv` — latest, 5-year, 52-week and 2025 average diesel price per region.
- `data/state_diesel_region.csv` — state → PADD → diesel price series.
- `scripts/tractor_model.py` — the shared tractor cost model used by both workbooks.
- `data/cattle_2025.csv.gz` — filtered 2025 auction observation rows (the extract).
- `data/report_inventory.json` — discovered USDA reports and category combinations.
