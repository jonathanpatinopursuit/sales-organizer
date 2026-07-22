#!/usr/bin/env python3
"""Ask a plain-language question about your sales data.

Usage:
    python3 scripts/ask.py "why is the West region down this month?"
    python3 scripts/ask.py "how is the Electronics category doing?"
    python3 scripts/ask.py "what's going on with Widget Pro?"

This is rule-based (no external API/key needed): it looks for a region,
category, or product name from your data inside the question, then pulls
current-vs-prior-period numbers for that segment straight from data/.
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd

import analysis
import common


def _best_match(question: str, candidates: list[str]) -> str | None:
    q = question.lower()
    matches = [c for c in candidates if c and c.lower() in q]
    if not matches:
        return None
    # Prefer the longest match (most specific name) if several candidates overlap.
    return max(matches, key=len)


def _segment_report(data: pd.DataFrame, dim: str, name: str) -> str:
    current_df, prior_df, current_period, prior_period = common.split_periods(data)

    cur_seg = current_df[current_df[dim].str.lower() == name.lower()]
    pri_seg = prior_df[prior_df[dim].str.lower() == name.lower()]

    cur_rev = cur_seg["revenue"].sum()
    pri_rev = pri_seg["revenue"].sum() if not pri_seg.empty else None
    change = common.pct_change(cur_rev, pri_rev) if pri_rev else None
    cur_profit = cur_seg["profit"].sum()
    cur_margin = (cur_profit / cur_rev * 100) if cur_rev else None
    pri_margin = (pri_seg["profit"].sum() / pri_rev * 100) if pri_rev else None

    lines = [f"## {dim.title()}: {name}\n"]

    if cur_seg.empty:
        return f"## {dim.title()}: {name}\n\nNo rows found for '{name}' in the current period ({current_period})."

    if change is None:
        lines.append(f"{name} did ${cur_rev:,.0f} in revenue in {current_period}. "
                     f"No prior-period data is available yet for a comparison.")
    else:
        direction = "up" if change >= 0 else "down"
        lines.append(f"{name} is **{direction} {abs(change):.1f}%** in {current_period} "
                     f"(${cur_rev:,.0f}) vs {prior_period} (${pri_rev:,.0f}).")

    if cur_margin is not None:
        margin_note = f"Current margin is {cur_margin:.1f}%"
        if pri_margin is not None:
            margin_note += f" (was {pri_margin:.1f}% in {prior_period})."
        else:
            margin_note += "."
        lines.append(margin_note)

    # Break down the "other" dimensions within this segment to explain *why*.
    other_dims = [d for d in ("category", "region", "product") if d != dim]
    for other in other_dims:
        seg_cur = cur_seg
        seg_pri = pri_seg
        breakdown = analysis.grouped_summary(seg_cur, seg_pri, other)
        breakdown = breakdown.dropna(subset=["pct_change"])
        if breakdown.empty:
            continue
        worst = breakdown.sort_values("pct_change").head(3)
        declines = worst[worst["pct_change"] < 0]
        if not declines.empty:
            bits = ", ".join(f"{r[other]} ({r['pct_change']:.1f}%)" for _, r in declines.iterrows())
            lines.append(f"Within {name}, the biggest {other} declines were: {bits}.")

    # Discount / margin risk within this segment
    disc = analysis.discount_analysis(cur_seg, "product", top_n=3)
    risky = disc[disc["margin_risk"]] if not disc.empty else disc
    if not risky.empty:
        bits = ", ".join(f"{r['product']} ({r['avg_discount']*100:.0f}% avg discount, "
                          f"{r['margin']*100:.1f}% margin)" for _, r in risky.iterrows())
        lines.append(f"Heavy discounting may be hurting margin on: {bits}.")

    if change is not None and change <= analysis.DECLINE_FLAG_THRESHOLD:
        lines.append(f"\nThis crosses the underperforming threshold "
                     f"({analysis.DECLINE_FLAG_THRESHOLD:.0f}% revenue decline) used in the weekly report's Flags section.")

    return "\n\n".join(lines)


def _general_report(data: pd.DataFrame, question: str) -> str:
    current_df, prior_df, current_period, prior_period = common.split_periods(data)
    category_df = analysis.category_summary(current_df, prior_df)
    region_df = analysis.region_summary(current_df, prior_df)
    product_df = analysis.product_summary(current_df, prior_df)
    flags = analysis.generate_flags(category_df, region_df, product_df)

    summary = analysis.build_summary_paragraph(
        current_df, prior_df, current_period, prior_period, category_df, region_df, flags
    )

    known = sorted(set(data["region"]) | set(data["category"]) | set(data["product"]))
    hint = ", ".join(known[:12]) + ("…" if len(known) > 12 else "")

    return (f"I couldn't find a specific region, category, or product from your data in that question.\n\n"
            f"Here's the general picture instead:\n\n{summary}\n\n"
            f"Try naming a specific one, e.g. \"why is the West region down?\" "
            f"Known names include: {hint}")


def answer(question: str) -> str:
    data = common.load_data()

    region_match = _best_match(question, sorted(data["region"].unique(), key=len, reverse=True))
    category_match = _best_match(question, sorted(data["category"].unique(), key=len, reverse=True))
    product_match = _best_match(question, sorted(data["product"].unique(), key=len, reverse=True))

    # Most specific match wins: product > category > region.
    if product_match:
        return _segment_report(data, "product", product_match)
    if category_match:
        return _segment_report(data, "category", category_match)
    if region_match:
        return _segment_report(data, "region", region_match)
    return _general_report(data, question)


def main():
    if len(sys.argv) < 2:
        print('Usage: python3 scripts/ask.py "your question here"')
        sys.exit(1)
    question = " ".join(sys.argv[1:])
    print(answer(question))


if __name__ == "__main__":
    main()
