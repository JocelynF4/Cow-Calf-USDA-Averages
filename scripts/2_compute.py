import csv
import gzip
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "data", "cattle_2025.csv.gz")
OUT = os.path.join(ROOT, "data", "averages_long.csv")

CWT, HEAD = "Per Cwt", "Per Unit"

SPECS = [
    ("Feeder Steers", "Feeder Cattle", "Steers", CWT),
    ("Feeder Heifers", "Feeder Cattle", "Heifers", CWT),
    ("Feeder Bulls", "Feeder Cattle", "Bulls", CWT),
    ("Feeder Cows", "Feeder Cattle", "Cows", CWT),
    ("Stock Cows ($/cwt)", "Replacement Cattle", "Stock Cows", CWT),
    ("Stock Cows ($/head)", "Replacement Cattle", "Stock Cows", HEAD),
    ("Bred Cows ($/head)", "Replacement Cattle", "Bred Cows", HEAD),
    ("Bred Cows ($/cwt)", "Replacement Cattle", "Bred Cows", CWT),
    ("Bred Heifers ($/head)", "Replacement Cattle", "Bred Heifers", HEAD),
    ("Bred Heifers ($/cwt)", "Replacement Cattle", "Bred Heifers", CWT),
    ("Cow-Calf Pairs ($/head)", "Replacement Cattle", "Cow-Calf Pairs", HEAD),
    ("Heifer Pairs ($/head)", "Replacement Cattle", "Heifer Pairs", HEAD),
    ("Replacement Bulls ($/head)", "Replacement Cattle", "Bulls", HEAD),
    ("Replacement Bulls ($/cwt)", "Replacement Cattle", "Bulls", CWT),
]

OVERALL_CLASSES = {"Steers", "Heifers", "Bulls"}


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def weighted(subset):
    total, head, obs, weeks = 0.0, 0.0, 0, set()
    for r in subset:
        price, h = num(r["avg_price"]), num(r["head_count"]) or 0
        if price is None or h <= 0:
            continue
        total += price * h
        head += h
        obs += 1
        weeks.add(r["report_date"])
    return (total / head if head else None), int(head), obs, len(weeks)


def main():
    with gzip.open(SRC, "rt") as fh:
        rows = list(csv.DictReader(fh))
    states = sorted({r["state"] for r in rows})

    with open(OUT, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["state", "category", "unit", "avg_price", "total_head",
                    "n_obs", "n_weeks"])
        for state in states:
            srows = [r for r in rows if r["state"] == state]

            ov = weighted([r for r in srows
                           if r["commodity"] == "Feeder Cattle"
                           and r["class"] in OVERALL_CLASSES
                           and r["price_unit"] == CWT])
            if ov[0] is not None:
                w.writerow([state, "Overall Feeder", "$/cwt", round(ov[0], 2),
                            ov[1], ov[2], ov[3]])

            for label, commodity, cls, unit in SPECS:
                sub = [r for r in srows if r["commodity"] == commodity
                       and r["class"] == cls and r["price_unit"] == unit]
                avg, head, obs, weeks = weighted(sub)
                if avg is None:
                    continue
                w.writerow([state, label,
                            "$/cwt" if unit == CWT else "$/head",
                            round(avg, 2), head, obs, weeks])
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
