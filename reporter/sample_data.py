"""Synthetic monthly branch sales used to demo the report."""
from __future__ import annotations

import numpy as np
import pandas as pd


def sample_sales(month: str = "2025-11", seed: int = 4) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    days = pd.date_range(f"{month}-01", periods=pd.Period(month).days_in_month, freq="D")
    branches = ["Karachi", "Lahore", "Islamabad", "Multan", "Peshawar"]
    cats = ["Mobiles", "Laptops", "Audio", "Accessories"]
    rows = [{"date": d.date(), "branch": b, "category": c,
             "units": int(rng.poisson(20 if c == "Accessories" else 6)),
             "revenue": float(round(rng.gamma(4, {"Mobiles": 15000, "Laptops": 40000, "Audio": 3000,
                                                    "Accessories": 600}[c]), 0)),
             "target": {"Mobiles": 60000, "Laptops": 150000, "Audio": 12000, "Accessories": 2500}[c]}
            for d in days for b in branches for c in cats]
    return pd.DataFrame(rows)
