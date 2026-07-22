"""Shared helpers for loading and normalizing weekly sales exports."""

from __future__ import annotations

import glob
import os

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

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")


def find_data_files(data_dir: str = DATA_DIR) -> list[str]:
    patterns = ["*.xlsx", "*.xls"]
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(data_dir, pattern)))
    # Ignore Excel lock files like ~$export.xlsx
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    return sorted(files)


def load_data(data_dir: str = DATA_DIR) -> pd.DataFrame:
    """Load and concatenate every Excel export found in data_dir."""
    files = find_data_files(data_dir)
    if not files:
        raise FileNotFoundError(
            f"No Excel files found in {data_dir}. Drop your weekly export "
            f"(.xlsx) there and run this script again."
        )

    frames = []
    for f in files:
        df = pd.read_excel(f)
        df.columns = [str(c).strip().lower() for c in df.columns]
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(
                f"{os.path.basename(f)} is missing expected column(s): {', '.join(missing)}. "
                f"Expected columns: {', '.join(REQUIRED_COLUMNS)}"
            )
        df["__source_file"] = os.path.basename(f)
        frames.append(df[REQUIRED_COLUMNS + ["__source_file"]])

    data = pd.concat(frames, ignore_index=True)

    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    bad_dates = data["date"].isna().sum()
    if bad_dates:
        data = data.dropna(subset=["date"])

    for col in ("quantity", "price", "discount", "profit"):
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    for col in ("customer", "product", "category", "region"):
        data[col] = data[col].astype(str).str.strip()

    # Revenue = quantity * price * (1 - discount), where discount is a 0..1 rate.
    # If discount looks like it's stored as a percentage (e.g. 10 instead of 0.10),
    # normalize it down to a rate.
    disc = data["discount"]
    if (disc > 1).any():
        data["discount"] = disc.where(disc <= 1, disc / 100.0)

    data["revenue"] = data["quantity"] * data["price"] * (1 - data["discount"])
    data["margin"] = (data["profit"] / data["revenue"].replace(0, pd.NA)).astype(float)
    data["period"] = data["date"].dt.to_period("M")

    return data.sort_values("date").reset_index(drop=True)


def current_and_prior_period(data: pd.DataFrame):
    """Return (current_period, prior_period) as pandas Period objects, based on
    the most recent two calendar months present in the data. If only one month
    is present, prior_period is None."""
    periods = sorted(data["period"].dropna().unique())
    if not periods:
        return None, None
    current = periods[-1]
    prior = periods[-2] if len(periods) > 1 else None
    return current, prior


def split_periods(data: pd.DataFrame):
    current, prior = current_and_prior_period(data)
    current_df = data[data["period"] == current] if current is not None else data.iloc[0:0]
    prior_df = data[data["period"] == prior] if prior is not None else data.iloc[0:0]
    return current_df, prior_df, current, prior


def pct_change(current: float, prior: float):
    if prior == 0 or pd.isna(prior):
        return None
    return (current - prior) / abs(prior) * 100.0
