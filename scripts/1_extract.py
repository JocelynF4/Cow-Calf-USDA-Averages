import csv
import gzip
import json
import os
import subprocess
import sys

API_KEY = os.environ.get("AMS_API_KEY")
if not API_KEY:
    sys.exit("Set AMS_API_KEY in the environment (USDA Market News / MARS API key).")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "cattle_2025.csv.gz")
START, END = "01/01/2025", "12/31/2025"

STATES = {
    "Alabama": 2006, "Arkansas": 2056, "Colorado": 1907, "Florida": 1704,
    "Georgia": 1933, "Illinois": 2041, "Iowa": 2167, "Kansas": 1895,
    "Kentucky": 2193, "Mississippi": 2115, "Missouri": 1821, "Montana": 1778,
    "Nebraska": 1860, "New York": 2011, "North Carolina": 2091,
    "North Dakota": 2100, "Oklahoma": 1831, "Pennsylvania": 1919,
    "South Carolina": 1963, "South Dakota": 2027, "Tennessee": 2063,
    "Texas": 1955, "Utah": 2039, "Virginia": 2187, "West Virginia": 1868,
    "Wyoming": 2106,
}

KEEP_COMMODITY = {"Feeder Cattle", "Replacement Cattle"}
FIELDS = ["state", "report_date", "commodity", "class", "price_unit", "avg_price",
          "head_count", "avg_price_min", "avg_price_max", "avg_weight",
          "wt_low", "wt_high"]


def is_dairy(commodity, cls):
    return "dairy" in (str(commodity) + " " + str(cls)).lower()


def fetch(slug):
    url = (f"https://marsapi.ams.usda.gov/services/v1.2/reports/{slug}"
           f"?q=report_date={START}:{END}")
    r = subprocess.run(
        ["curl", "-sS", "-u", f"{API_KEY}:", "--max-time", "300", url],
        capture_output=True, text=True,
    )
    return json.loads(r.stdout).get("results", [])


def main():
    kept = 0
    with gzip.open(OUT, "wt", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(FIELDS)
        for state, slug in sorted(STATES.items()):
            rows = fetch(slug)
            n = 0
            for row in rows:
                commodity, cls = row.get("commodity"), row.get("class")
                if commodity not in KEEP_COMMODITY or is_dairy(commodity, cls):
                    continue
                try:
                    price = float(row.get("avg_price"))
                except (TypeError, ValueError):
                    continue
                w.writerow([
                    state, row.get("report_date"), commodity, cls,
                    row.get("price_unit"), price, row.get("head_count"),
                    row.get("avg_price_min"), row.get("avg_price_max"),
                    row.get("avg_weight"), row.get("weight_break_low"),
                    row.get("weight_break_high"),
                ])
                n += 1
            kept += n
            print(f"{state:<16} slug={slug} kept={n}")
    print(f"wrote {kept} rows to {OUT}")


if __name__ == "__main__":
    main()
