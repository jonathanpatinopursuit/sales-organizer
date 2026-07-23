"""Shared helpers for loading and normalizing weekly sales exports."""

from __future__ import annotations

import glob
import os

import pandas as pd

from validate_data import REQUIRED_COLUMNS, validate

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")


def find_data_files(data_dir: str = DATA_DIR) -> list[str]:
    patterns = ["*.xlsx", "*.xls"]
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(data_dir, pattern)))
    # Ignore Excel lock files like ~$export.xlsx
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    return sorted(files)


EMPTY_COLUMNS = REQUIRED_COLUMNS + ["__source_file", "revenue", "margin", "period"]


def load_data(data_dir: str = DATA_DIR) -> tuple[pd.DataFrame, list[dict], list[str]]:
    """Load, validate, and normalize every Excel export found in data_dir.

    Returns (data, issues, halts):
      - issues: non-fatal data-quality dicts from validate_data.validate()
        (skipped rows, clamped discounts, duplicates) for the report's banner.
      - halts: messages for files that were rejected outright (missing
        required columns, or too many unparseable dates). A halted file is
        excluded and the run continues with whatever files remain valid —
        it does not crash the whole run; the halt is surfaced in the report
        instead.
    """
    files = find_data_files(data_dir)
    if not files:
        raise FileNotFoundError(
            f"No Excel files found in {data_dir}. Drop your weekly export "
            f"(.xlsx) there and run this script again."
        )

    frames = []
    issues: list[dict] = []
    halts: list[str] = []
    for f in files:
        filename = os.path.basename(f)
        df = pd.read_excel(f)
        df.columns = [str(c).strip().lower() for c in df.columns]
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            halts.append(
                f"{filename} is missing expected column(s): {', '.join(missing)}. "
                f"Expected columns: {', '.join(REQUIRED_COLUMNS)}"
            )
            continue
        df = df[REQUIRED_COLUMNS].copy()

        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        for col in ("quantity", "price", "discount", "profit"):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Revenue = quantity * price * (1 - discount), where discount is a 0..1 rate.
        # If discount looks like it's stored as a percentage (e.g. 10 instead of
        # 0.10), normalize it down to a rate *before* validate() checks the range,
        # so a legitimate "10" doesn't get mistaken for a 1000% discount and
        # clamped to 100%.
        if (df["discount"] > 1).any():
            df["discount"] = df["discount"].where(df["discount"] <= 1, df["discount"] / 100.0)

        try:
            df, file_issues = validate(df, filename)
        except ValueError as e:
            halts.append(str(e))
            continue
        issues.extend(file_issues)

        df["__source_file"] = filename
        frames.append(df)

    if not frames:
        return pd.DataFrame(columns=EMPTY_COLUMNS), issues, halts

    data = pd.concat(frames, ignore_index=True)

    for col in ("customer", "product", "category", "region"):
        data[col] = data[col].astype(str).str.strip()

    # validate() only catches duplicates within a single file; also check
    # across files, in case the same export got saved under two filenames.
    cross_file_dupe_mask = data.duplicated(subset=REQUIRED_COLUMNS, keep=False)
    cross_file_dupes = int(data.duplicated(subset=REQUIRED_COLUMNS).sum())
    if cross_file_dupes:
        issues.append({
            "level": "warn",
            "message": (
                f"Found {cross_file_dupes} row(s) duplicated across multiple files in data/ "
                f"— check for the same export saved under more than one filename."
            ),
            "count": cross_file_dupes,
            "products": sorted(set(data.loc[cross_file_dupe_mask, "product"])),
            "categories": sorted(set(data.loc[cross_file_dupe_mask, "category"])),
        })

    data["revenue"] = data["quantity"] * data["price"] * (1 - data["discount"])
    data["margin"] = (data["profit"] / data["revenue"].replace(0, pd.NA)).astype(float)
    data["period"] = data["date"].dt.to_period("M")

    return data.sort_values("date").reset_index(drop=True), issues, halts


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
