"""Intake validation for weekly sales exports.

Called per-file from common.load_data(), after that file's date column has
been parsed (pd.to_datetime(errors="coerce")) and its numeric columns coerced
and discount normalized to a 0-1 rate — validate() assumes that's already
done, so its own checks are catching genuine bad data, not raw formatting.

Halting problems (missing columns, too many bad dates) raise ValueError.
common.load_data() catches that, excludes just that file, and keeps going
with whatever files remain rather than crashing the whole run — the halt
still shows up as a HALT banner in the generated report.

Non-halting problems are returned as a list of issue dicts:
    {"level": "skip" | "warn", "message": str, "count": int,
     "products": [...], "categories": [...]}
"products"/"categories" name which rows were affected (when applicable), so
the report can flag those specific rows inline, not just show a raw count.
"""

from __future__ import annotations

import pandas as pd

REQUIRED_COLUMNS = [
    "date",
    "customer",
    "product",
    "category",
    "region",
    "quantity",
    "price",
    "discount",
    "profit",
]

DATE_HALT_THRESHOLD = 0.05  # halt the file if more than this fraction of dates are bad


def validate(df: pd.DataFrame, filename: str = "input") -> tuple[pd.DataFrame, list[dict]]:
    issues: list[dict] = []

    # Check 1: missing required columns -- halts
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s) in {filename}: {missing}")

    # Check 2: unparseable/blank dates (already NaT by the time this runs) -- halts above threshold
    bad_dates = df["date"].isna()
    pct = bad_dates.mean() if len(df) else 0.0
    if pct > DATE_HALT_THRESHOLD:
        raise ValueError(
            f"{pct:.1%} of dates unparseable in {filename} — likely wrong column/format."
        )
    elif bad_dates.any():
        issues.append({
            "level": "skip",
            "message": f"Skipped {bad_dates.sum()} row(s) with unparseable/blank dates in {filename} ({pct:.1%}).",
            "count": int(bad_dates.sum()),
        })
        df = df[~bad_dates]

    # Check 3: non-positive quantity or negative price
    bad_qty = (df["quantity"] <= 0) | (df["price"] < 0)
    if bad_qty.any():
        issues.append({
            "level": "skip",
            "message": f"Skipped {bad_qty.sum()} row(s) with non-positive quantity or negative price in {filename}.",
            "count": int(bad_qty.sum()),
        })
        df = df[~bad_qty]

    # Check 4: discount outside 0-100% (post-normalization, so this really is bad data)
    bad_disc = (df["discount"] < 0) | (df["discount"] > 1)
    if bad_disc.any():
        issues.append({
            "level": "warn",
            "message": f"Clamped {bad_disc.sum()} discount value(s) to [0, 1] in {filename}.",
            "count": int(bad_disc.sum()),
            "products": sorted(set(df.loc[bad_disc, "product"])),
            "categories": sorted(set(df.loc[bad_disc, "category"])),
        })
        df = df.copy()
        df["discount"] = df["discount"].clip(0, 1)

    # Check 5: duplicate rows within this file (not removed -- legit repeat orders can look identical)
    dupe_mask = df.duplicated(keep=False)
    dupes = int(df.duplicated().sum())
    if dupes:
        issues.append({
            "level": "warn",
            "message": f"Found {dupes} duplicate row(s) within {filename} — not removed.",
            "count": dupes,
            "products": sorted(set(df.loc[dupe_mask, "product"])),
            "categories": sorted(set(df.loc[dupe_mask, "category"])),
        })

    return df, issues
