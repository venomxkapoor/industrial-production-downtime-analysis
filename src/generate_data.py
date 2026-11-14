"""
Generate a reproducible set of SIMULATED production orders.

The data is fictional. It is created with Python's random module and a
fixed seed, so every run produces exactly the same CSV file.

Run from the project folder:
    python src/generate_data.py
"""

import random
from pathlib import Path

import pandas as pd

SEED = 42
NUMBER_OF_ORDERS = 80

# Two independent work centres. About 60 % of orders are assigned to WC1.
# This uneven split is a scenario assumption chosen before looking at results.
WORK_CENTRES = ["WC1", "WC2"]
SHARE_OF_ORDERS_FOR_WC1 = 0.6

PRODUCT_CATEGORIES = ["Bracket", "Housing", "Shaft", "Panel"]

# Processing time: 1.0 to 8.0 hours in steps of 0.5 hours.
# We draw a whole number of half-hours (2 to 16) and divide by 2.
MIN_HALF_HOURS = 2
MAX_HALF_HOURS = 16

# Due time: a whole number of hours after the start of the planning period.
EARLIEST_DUE_TIME_H = 8
LATEST_DUE_TIME_H = 160

# All orders are available at time zero.
RELEASE_TIME_H = 0

# Path of the project folder (the folder that contains "src" and "data").
PROJECT_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_FILE = PROJECT_FOLDER / "data" / "production_orders.csv"


def generate_orders(number_of_orders, seed):
    """Create a DataFrame with simulated production orders."""
    random.seed(seed)
    rows = []

    for sequence in range(1, number_of_orders + 1):
        if random.random() < SHARE_OF_ORDERS_FOR_WC1:
            work_centre = WORK_CENTRES[0]
        else:
            work_centre = WORK_CENTRES[1]

        category = random.choice(PRODUCT_CATEGORIES)
        processing_time = random.randint(MIN_HALF_HOURS, MAX_HALF_HOURS) / 2
        due_time = random.randint(EARLIEST_DUE_TIME_H, LATEST_DUE_TIME_H)

        row = {
            "order_id": f"ORD-{sequence:03d}",
            "original_sequence": sequence,
            "work_centre": work_centre,
            "product_category": category,
            "processing_time_h": processing_time,
            "due_time_h": due_time,
            "release_time_h": RELEASE_TIME_H,
        }
        rows.append(row)

    return pd.DataFrame(rows)


def main():
    orders = generate_orders(NUMBER_OF_ORDERS, SEED)
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    orders.to_csv(OUTPUT_FILE, index=False)
    print(f"Saved {len(orders)} simulated orders to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
