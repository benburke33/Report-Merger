#!/usr/bin/env python3
"""

Usage:
    python sample_reports.py                # writes into ./weekly_reports/
    python sample_reports.py --outdir demo --seed 7

Creates week1.csv, week2.csv, week3.xlsx, week4.csv with the same columns:
region, product, units, revenue. Week 3 is an Excel file and week 4 stores revenue
as text ("$1,234.50") so the merger's mixed-format handling gets exercised.
All data is invented.
"""
import argparse
import random
from pathlib import Path

import pandas as pd

REGIONS = ["West", "East", "North", "South"]
PRODUCTS = {"Lavender Candle": 18.0, "Cedar Candle": 22.0, "Soy Wax Melts": 9.5, "Gift Set": 45.0}


def make_week(rng: random.Random):
    rows = []
    for region in REGIONS:
        for product, price in PRODUCTS.items():
            if rng.random() < 0.15:          # not every region sells every product every week
                continue
            units = rng.randint(8, 70)
            rows.append({"region": region, "product": product, "units": units,
                         "revenue": round(units * price * rng.uniform(0.92, 1.0), 2)})
    return pd.DataFrame(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--outdir", type=Path, default=Path("weekly_reports"))
    ap.add_argument("--seed", type=int, default=42)
    args = ap.parse_args()

    rng = random.Random(args.seed)
    args.outdir.mkdir(parents=True, exist_ok=True)

    for n in range(1, 5):
        df = make_week(rng)
        if n == 3:
            df.to_excel(args.outdir / f"week{n}.xlsx", index=False)
        else:
            if n == 4:
                df["revenue"] = df["revenue"].map(lambda v: f"${v:,.2f}")
            df.to_csv(args.outdir / f"week{n}.csv", index=False)
    print(f"Wrote 4 weekly files to {args.outdir}/")


if __name__ == "__main__":
    main()