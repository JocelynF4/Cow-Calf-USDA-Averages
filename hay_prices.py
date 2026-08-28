import csv
import os
from collections import defaultdict

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Hay prices state June-Sept.csv")

NO_DATA_MESSAGE = 'no hay price data available for {state}, manual entry required'


def _load_state_averages(data_file=DATA_FILE):
    values = defaultdict(list)
    with open(data_file, newline="") as f:
        for row in csv.DictReader(f):
            if row["Geo Level"].strip().upper() != "STATE":
                continue
            if "PRICE RECEIVED" not in row["Data Item"].upper():
                continue
            raw = row["Value"].strip().replace(",", "")
            if not raw or not raw.replace(".", "", 1).isdigit():
                continue
            values[row["State"].strip().upper()].append(float(raw))
    return {state: sum(v) / len(v) for state, v in values.items()}


STATE_AVERAGES = _load_state_averages()


def get_hay_price(county, state):
    key = (state or "").strip().upper()
    if key in STATE_AVERAGES:
        return round(STATE_AVERAGES[key], 2)
    return NO_DATA_MESSAGE.format(state=(state or "").strip())


if __name__ == "__main__":
    county = input("County: ")
    state = input("State: ")
    result = get_hay_price(county, state)
    if isinstance(result, float):
        print(f"Average hay price for {state.strip().title()}: ${result:.2f}/ton")
    else:
        print(result)
