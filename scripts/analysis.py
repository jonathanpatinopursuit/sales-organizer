"""Core analysis functions shared by the report generator and the Q&A tool."""

from __future__ import annotations

import pandas as pd

from common import pct_change

DECLINE_FLAG_THRESHOLD = -15.0   # % revenue drop vs prior period
LOW_MARGIN_THRESHOLD = 0.10      # margin below 10% is flagged
HIGH_DISCOUNT_THRESHOLD = 0.15   # average discount above 15% is "big"
NEGATIVE_PROFIT_FLAG = True      # always flag any segment with negative total profit


def grouped_summary(current_df: pd.DataFrame, prior_df: pd.DataFrame, by: str) -> pd.DataFrame:
    def agg(df):
        if df.empty:
            return pd.DataFrame(columns=[by, "revenue", "profit", "quantity", "avg_discount"])
        g = df.groupby(by).agg(
            revenue=("revenue", "sum"),
            profit=("profit", "sum"),
            quantity=("quantity", "sum"),
            avg_discount=("discount", "mean"),
        ).reset_index()
        return g

    cur = agg(current_df)
    pri = agg(prior_df)

    merged = cur.merge(pri[[by, "revenue"]].rename(columns={"revenue": "prior_revenue"}), on=by, how="left")
    merged["margin"] = (merged["profit"] / merged["revenue"].replace(0, pd.NA)).astype(float)
    merged["pct_change"] = merged.apply(
        lambda r: pct_change(r["revenue"], r["prior_revenue"]) if pd.notna(r.get("prior_revenue")) else None,
        axis=1,
    )
    return merged.sort_values("revenue", ascending=False).reset_index(drop=True)


def category_summary(current_df: pd.DataFrame, prior_df: pd.DataFrame) -> pd.DataFrame:
    return grouped_summary(current_df, prior_df, "category")


def region_summary(current_df: pd.DataFrame, prior_df: pd.DataFrame) -> pd.DataFrame:
    return grouped_summary(current_df, prior_df, "region")


def product_summary(current_df: pd.DataFrame, prior_df: pd.DataFrame) -> pd.DataFrame:
    return grouped_summary(current_df, prior_df, "product")


def discount_analysis(current_df: pd.DataFrame, group_col: str = "product", top_n: int = 10) -> pd.DataFrame:
    if current_df.empty:
        return pd.DataFrame(columns=[group_col, "avg_discount", "revenue", "profit", "margin", "margin_risk"])
    g = current_df.groupby(group_col).agg(
        avg_discount=("discount", "mean"),
        revenue=("revenue", "sum"),
        profit=("profit", "sum"),
    ).reset_index()
    g["margin"] = (g["profit"] / g["revenue"].replace(0, pd.NA)).astype(float)
    g["margin_risk"] = (g["avg_discount"] >= HIGH_DISCOUNT_THRESHOLD) & (g["margin"] < LOW_MARGIN_THRESHOLD)
    return g.sort_values("avg_discount", ascending=False).head(top_n).reset_index(drop=True)


def generate_flags(category_df: pd.DataFrame, region_df: pd.DataFrame, product_df: pd.DataFrame) -> list[dict]:
    flags = []

    def scan(df, dim_label):
        for _, row in df.iterrows():
            name = row[dim_label if dim_label in df.columns else df.columns[0]]
            reasons = []
            severity = "warning"
            if pd.notna(row.get("pct_change")) and row["pct_change"] <= DECLINE_FLAG_THRESHOLD:
                reasons.append(f"revenue down {row['pct_change']:.1f}% vs prior period")
                severity = "critical" if row["pct_change"] <= 2 * DECLINE_FLAG_THRESHOLD else "warning"
            if pd.notna(row.get("margin")) and row["margin"] < LOW_MARGIN_THRESHOLD:
                reasons.append(f"margin only {row['margin'] * 100:.1f}%")
            if NEGATIVE_PROFIT_FLAG and row.get("profit", 0) < 0:
                reasons.append(f"negative total profit (${row['profit']:,.0f})")
                severity = "critical"
            if reasons:
                flags.append({
                    "dimension": dim_label,
                    "name": name,
                    "reason": "; ".join(reasons),
                    "severity": severity,
                })

    scan(category_df.rename(columns={"category": "category"}), "category")
    scan(region_df.rename(columns={"region": "region"}), "region")
    scan(product_df.rename(columns={"product": "product"}), "product")

    severity_order = {"critical": 0, "warning": 1}
    flags.sort(key=lambda f: severity_order.get(f["severity"], 2))
    return flags


def build_summary_paragraph(current_df, prior_df, current_period, prior_period,
                             category_df, region_df, flags) -> str:
    total_revenue = current_df["revenue"].sum()
    total_profit = current_df["profit"].sum()
    overall_margin = (total_profit / total_revenue * 100) if total_revenue else 0

    period_label = str(current_period) if current_period is not None else "this period"

    parts = [
        f"In {period_label}, total sales revenue was ${total_revenue:,.0f} "
        f"with ${total_profit:,.0f} in profit (a {overall_margin:.1f}% overall margin)."
    ]

    if prior_period is not None and not prior_df.empty:
        prior_revenue = prior_df["revenue"].sum()
        change = pct_change(total_revenue, prior_revenue)
        if change is not None:
            direction = "up" if change >= 0 else "down"
            parts.append(f"That's {direction} {abs(change):.1f}% versus {prior_period} (${prior_revenue:,.0f}).")
    else:
        parts.append("No prior period was found in the data yet, so period-over-period comparisons aren't available.")

    if not category_df.empty:
        top_cat = category_df.iloc[0]
        parts.append(f"{top_cat['category']} was the top category by revenue (${top_cat['revenue']:,.0f}).")

    if not region_df.empty and region_df["pct_change"].notna().any():
        worst = region_df.dropna(subset=["pct_change"]).sort_values("pct_change").iloc[0]
        best = region_df.dropna(subset=["pct_change"]).sort_values("pct_change").iloc[-1]
        if worst["pct_change"] < 0:
            parts.append(f"{worst['region']} was the weakest region ({worst['pct_change']:.1f}%), "
                         f"while {best['region']} led growth ({best['pct_change']:+.1f}%).")

    n_critical = sum(1 for f in flags if f["severity"] == "critical")
    n_warning = sum(1 for f in flags if f["severity"] == "warning")
    if n_critical or n_warning:
        parts.append(f"{n_critical} critical and {n_warning} warning flag(s) were raised below — see Flags for detail.")
    else:
        parts.append("No critical issues were flagged this period.")

    return " ".join(parts)
