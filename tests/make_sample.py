"""Generate a realistic sample workbook for manual/automated testing."""
from __future__ import annotations

import datetime as _dt
import os
import random

from openpyxl import Workbook


def make_sample(path: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Data"
    ws.append(["Date", "Region", "Product", "Category", "Units", "Revenue", "Margin %"])

    regions = ["North", "South", "East", "West"]
    products = ["Widget", "Gadget", "Gizmo", "Doohickey"]
    categories = ["Hardware", "Software", "Services"]
    random.seed(42)
    start = _dt.date(2024, 1, 1)
    for i in range(400):
        d = start + _dt.timedelta(days=random.randint(0, 364))
        units = random.randint(1, 50)
        price = random.uniform(10, 500)
        revenue = round(units * price, 2)
        ws.append([
            _dt.datetime(d.year, d.month, d.day),
            random.choice(regions), random.choice(products),
            random.choice(categories), units, revenue,
            round(random.uniform(0.05, 0.45), 3),
        ])

    # A second sheet to exercise multi-sheet auto-detect.
    ws2 = wb.create_sheet("Targets")
    ws2.append(["Region", "Target Revenue"])
    for r in regions:
        ws2.append([r, random.randint(50000, 150000)])

    wb.save(path)
    return path


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(__file__), "sample.xlsx")
    print("Wrote", make_sample(out))
