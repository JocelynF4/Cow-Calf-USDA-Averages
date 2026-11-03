"""Average the weekly on-highway diesel prices by fuel region and map every state to one.

Input : Weekly_On-Highway_Diesel_Fuel_Prices_20261103.csv  (weekly $/gal by region)
Output: data/diesel_by_region.csv   region -> latest week, 5-year, 52-week and 2025 averages
        data/state_diesel_region.csv state  -> fuel region (PADD) + which price series it uses
"""
import csv
import os
from collections import defaultdict
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "Weekly_On-Highway_Diesel_Fuel_Prices_20261103.csv")
OUT_REGION = os.path.join(ROOT, "data", "diesel_by_region.csv")
OUT_STATE = os.path.join(ROOT, "data", "state_diesel_region.csv")

# Petroleum Administration for Defense Districts (PADD) -> states.
# PADD 1 is split into its three sub-districts, each of which has its own price series.
# PADD 5 uses the "California" series for California and "West Coast exc California"
# for every other PADD 5 state, since the source publishes both.
# PADD 6 and 7 (territories) have no price series in the source.
PADD_STATES = [
    ("PADD 1A", "New England", "New England",
     ["Connecticut", "Maine", "Massachusetts", "New Hampshire", "Rhode Island", "Vermont"]),
    ("PADD 1B", "Central Atlantic", "Central Atlantic",
     ["Delaware", "District of Columbia", "Maryland", "New Jersey", "New York", "Pennsylvania"]),
    ("PADD 1C", "Lower Atlantic", "Lower Atlantic",
     ["Florida", "Georgia", "North Carolina", "South Carolina", "Virginia", "West Virginia"]),
    ("PADD 2", "Midwest", "Midwest",
     ["Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Michigan", "Minnesota", "Missouri",
      "Nebraska", "North Dakota", "Ohio", "Oklahoma", "South Dakota", "Tennessee", "Wisconsin"]),
    ("PADD 3", "Gulf Coast", "Gulf Coast",
     ["Alabama", "Arkansas", "Louisiana", "Mississippi", "New Mexico", "Texas"]),
    ("PADD 4", "Rocky Mountain", "Rocky Mountains",
     ["Colorado", "Idaho", "Montana", "Utah", "Wyoming"]),
    ("PADD 5", "West Coast", "California", ["California"]),
    ("PADD 5", "West Coast", "West Coast exc California",
     ["Alaska", "Arizona", "Hawaii", "Nevada", "Oregon", "Washington"]),
    ("PADD 6", "Caribbean territories", "", ["Puerto Rico", "U.S. Virgin Islands"]),
    ("PADD 7", "Pacific territories", "",
     ["Guam", "American Samoa", "Northern Mariana Islands"]),
]

WINDOW_WEEKS = 52
FIVE_YEAR_WEEKS = 260
CALENDAR_YEAR = 2025


def load_weekly(path=SRC):
    series = defaultdict(dict)
    with open(path, newline="") as fh:
        for r in csv.DictReader(fh):
            date = datetime.strptime(r["Date"].strip(), "%m/%d/%Y").date()
            raw = r["Diesel Price"].strip().replace("$", "").replace(",", "")
            if not raw:
                continue
            series[r["Region"].strip()][date] = float(raw)
    return series


def summarise(series):
    out = []
    for region, pts in series.items():
        dates = sorted(pts)
        last = dates[-1]
        window = dates[-WINDOW_WEEKS:]
        five = dates[-FIVE_YEAR_WEEKS:]
        if len(five) < FIVE_YEAR_WEEKS:
            raise SystemExit(f"{region}: only {len(five)} weeks, need {FIVE_YEAR_WEEKS} for 5 years")
        year = [d for d in dates if d.year == CALENDAR_YEAR]
        out.append({
            "region": region,
            "latest_week": last.isoformat(),
            "latest_price": round(pts[last], 3),
            "avg_5yr": round(sum(pts[d] for d in five) / len(five), 3),
            "avg_5yr_from": five[0].isoformat(),
            "avg_5yr_weeks": len(five),
            "avg_52wk": round(sum(pts[d] for d in window) / len(window), 3),
            "avg_52wk_from": window[0].isoformat(),
            "avg_52wk_weeks": len(window),
            f"avg_{CALENDAR_YEAR}": round(sum(pts[d] for d in year) / len(year), 3),
            f"avg_{CALENDAR_YEAR}_weeks": len(year),
        })
    return sorted(out, key=lambda r: r["region"])


def main():
    series = load_weekly()
    rows = summarise(series)
    with open(OUT_REGION, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)

    known = {r["region"] for r in rows}
    with open(OUT_STATE, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["state", "padd", "padd_name", "price_series"])
        for padd, name, serie, states in PADD_STATES:
            if serie and serie not in known:
                raise SystemExit(f"price series {serie!r} missing from {SRC}")
            for st in states:
                w.writerow([st, padd, name, serie])

    for r in rows:
        print(f"{r['region']:<28} latest {r['latest_price']:.3f}  5yr {r['avg_5yr']:.3f}  "
              f"52wk {r['avg_52wk']:.3f}  {CALENDAR_YEAR} {r[f'avg_{CALENDAR_YEAR}']:.3f}")
    print("saved:", OUT_REGION, OUT_STATE)


if __name__ == "__main__":
    main()
