"""Load the AAA 2026 average diesel price for each state.

Input : AAA_State_Diesel_2026.csv  (state, 2026 average on-road diesel $/gal, from
        AAA state gas price averages: https://gasprices.aaa.com/state-gas-price-averages/)
Output: data/diesel_by_state.csv   state -> diesel_2026_avg
"""
import csv
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "AAA_State_Diesel_2026.csv")
OUT = os.path.join(ROOT, "data", "diesel_by_state.csv")
COUNTIES = os.path.join(ROOT, "data", "counties.csv")


def main():
    rows = []
    with open(SRC, newline="") as fh:
        for r in csv.DictReader(fh):
            state = r["State"].strip()
            raw = r["Diesel 2026 Avg ($/gal)"].strip().replace("$", "")
            if not state or not raw:
                continue
            rows.append((state, round(float(raw), 4)))
    rows.sort()

    with open(COUNTIES) as fh:
        county_states = {r["state_name"] for r in csv.DictReader(fh)}
    have = {s for s, _ in rows}
    missing = sorted(county_states - have)
    if missing:
        raise SystemExit(f"no AAA diesel price for: {', '.join(missing)}")

    with open(OUT, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["state", "diesel_2026_avg"])
        w.writerows(rows)
    prices = [p for _, p in rows]
    print(f"{len(rows)} states; low {min(prices):.3f}, high {max(prices):.3f}, "
          f"simple mean {sum(prices) / len(prices):.3f}")
    print("saved:", OUT)


if __name__ == "__main__":
    main()
