import json
import os
import re
import subprocess
import sys
from collections import defaultdict

API_KEY = os.environ.get("AMS_API_KEY")
if not API_KEY:
    sys.exit("Set AMS_API_KEY in the environment (USDA Market News / MARS API key).")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "report_inventory.json")
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


def api(path):
    url = f"https://marsapi.ams.usda.gov/services/v1.2/{path}"
    r = subprocess.run(["curl", "-sS", "-u", f"{API_KEY}:", "--max-time", "300", url],
                       capture_output=True, text=True)
    return json.loads(r.stdout)


def list_state_summaries():
    reports = api("reports")
    pat = re.compile(r"^([A-Z][A-Za-z ]+?) Weekly (Cattle|Livestock) Auction .*Summary$")
    out = []
    for r in reports:
        m = pat.match(r["report_title"])
        if m:
            out.append({"state": m.group(1).strip(), "slug_id": r["slug_id"],
                        "title": r["report_title"]})
    return sorted(out, key=lambda x: x["state"])


def inventory():
    combos = defaultdict(lambda: {"rows": 0, "priced_rows": 0, "head": 0, "states": set()})
    for state, slug in sorted(STATES.items()):
        data = api(f"reports/{slug}?q=report_date={START}:{END}")
        for row in data.get("results", []):
            key = (row.get("group"), row.get("category"), row.get("commodity"),
                   row.get("class"), row.get("price_unit"))
            c = combos[key]
            c["rows"] += 1
            c["states"].add(state)
            price = row.get("avg_price")
            if price not in (None, "", "None"):
                c["priced_rows"] += 1
                try:
                    c["head"] += int(float(row.get("head_count") or 0))
                except (TypeError, ValueError):
                    pass
    out = []
    for (group, category, commodity, cls, unit), c in combos.items():
        out.append({"group": group, "category": category, "commodity": commodity,
                    "class": cls, "price_unit": unit, "rows": c["rows"],
                    "priced_rows": c["priced_rows"], "total_head": c["head"],
                    "n_states": len(c["states"]), "states": sorted(c["states"])})
    out.sort(key=lambda x: -x["priced_rows"])
    return out


def main():
    result = {"state_summaries": list_state_summaries(), "combos": inventory()}
    json.dump(result, open(OUT, "w"), indent=2)
    print(f"wrote {OUT}: {len(result['state_summaries'])} state reports, "
          f"{len(result['combos'])} category combinations")


if __name__ == "__main__":
    main()
