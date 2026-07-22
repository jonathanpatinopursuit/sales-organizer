#!/usr/bin/env python3
"""Generate the Sales Organizer report from every Excel file in data/.

Usage:
    python3 scripts/generate_report.py

Outputs (written to reports/):
    sales_report_<current-period>.xlsx  - formatted workbook, one sheet per section
    sales_report_<current-period>.html  - single-file HTML report (view in a browser)
    latest.xlsx / latest.html           - always overwritten copies of the newest report
"""

from __future__ import annotations

import os
import shutil
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd

import analysis
import common

REPORTS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "reports")

# Palette (validated set, see dataviz skill / references/palette.md)
COLOR_INK_PRIMARY = "#0b0b0b"
COLOR_INK_SECONDARY = "#52514e"
COLOR_INK_MUTED = "#898781"
COLOR_SURFACE = "#fcfcfb"
COLOR_PAGE = "#f9f9f7"
COLOR_GRIDLINE = "#e1e0d9"
COLOR_BLUE = "#2a78d6"
COLOR_GOOD_TEXT = "#006300"
COLOR_CRITICAL = "#d03b3b"
COLOR_WARNING = "#fab219"
COLOR_SERIOUS = "#ec835a"
COLOR_GOOD = "#0ca30c"


# ---------------------------------------------------------------------------
# Excel workbook
# ---------------------------------------------------------------------------

def write_excel(path, summary_text, current_period, prior_period,
                 category_df, region_df, discount_product_df, discount_category_df, flags_df):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#0b0b0b", "font_color": "#ffffff",
            "border": 1, "align": "left",
        })
        # pct_change is already expressed in percentage points (e.g. 39.2 means +39.2%),
        # so use a literal "%" suffix rather than Excel's "%" format code, which would
        # multiply the value by 100 again.
        pct_fmt = workbook.add_format({"num_format": '+0.0"%";-0.0"%"'})
        money_fmt = workbook.add_format({"num_format": "$#,##0"})
        pct_plain_fmt = workbook.add_format({"num_format": "0.0%"})
        good_fmt = workbook.add_format({"font_color": COLOR_GOOD_TEXT, "bold": True})
        bad_fmt = workbook.add_format({"font_color": COLOR_CRITICAL, "bold": True})
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})

        # --- Summary sheet ---
        ws = workbook.add_worksheet("Summary")
        writer.sheets["Summary"] = ws
        ws.set_column("A:A", 100)
        ws.write("A1", "Sales Organizer Report", workbook.add_format({"bold": True, "font_size": 16}))
        ws.write("A2", f"Current period: {current_period}   |   Prior period: {prior_period}")
        ws.write("A4", summary_text, wrap_fmt)
        ws.set_row(3, 60)

        def write_table(df, sheet_name, pct_cols=(), money_cols=(), plain_pct_cols=()):
            df = df.copy()
            df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            ws = writer.sheets[sheet_name]
            for col_num, col_name in enumerate(df.columns):
                ws.write(1, col_num, col_name, header_fmt)
                width = max(14, len(str(col_name)) + 4)
                fmt = None
                if col_name in pct_cols:
                    fmt = pct_fmt
                elif col_name in plain_pct_cols:
                    fmt = pct_plain_fmt
                elif col_name in money_cols:
                    fmt = money_fmt
                ws.set_column(col_num, col_num, width, fmt)
            # conditional formatting on pct_change columns
            if "pct_change" in df.columns:
                col_idx = list(df.columns).index("pct_change")
                first_row, last_row = 2, len(df) + 1
                col_letter = chr(ord("A") + col_idx)
                ws.conditional_format(f"{col_letter}{first_row}:{col_letter}{last_row}",
                                       {"type": "cell", "criteria": "<", "value": 0, "format": bad_fmt})
                ws.conditional_format(f"{col_letter}{first_row}:{col_letter}{last_row}",
                                       {"type": "cell", "criteria": ">=", "value": 0, "format": good_fmt})
            return ws

        write_table(category_df, "By Category", pct_cols=["pct_change"],
                    money_cols=["revenue", "profit", "prior_revenue"], plain_pct_cols=["margin", "avg_discount"])
        write_table(region_df, "By Region", pct_cols=["pct_change"],
                    money_cols=["revenue", "profit", "prior_revenue"], plain_pct_cols=["margin", "avg_discount"])
        write_table(discount_product_df, "Discounts by Product", money_cols=["revenue", "profit"],
                    plain_pct_cols=["avg_discount", "margin"])
        write_table(discount_category_df, "Discounts by Category", money_cols=["revenue", "profit"],
                    plain_pct_cols=["avg_discount", "margin"])
        write_table(flags_df, "Flags")


# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------

def _fmt_money(v):
    return f"${v:,.0f}"


def _fmt_pct(v, signed=True):
    if v is None or pd.isna(v):
        return "n/a"
    return f"{v:+.1f}%" if signed else f"{v:.1f}%"


def _delta_span(v):
    if v is None or pd.isna(v):
        return '<span class="muted">n/a</span>'
    cls = "delta-up" if v >= 0 else "delta-down"
    arrow = "▲" if v >= 0 else "▼"
    return f'<span class="{cls}">{arrow} {abs(v):.1f}%</span>'


def _bar_row(label, value, max_value):
    pct = 0 if max_value == 0 else max(2, value / max_value * 100)
    return f"""
      <div class="bar-row">
        <div class="bar-label">{label}</div>
        <div class="bar-track"><div class="bar-fill" style="width:{pct:.1f}%"></div></div>
        <div class="bar-value">{_fmt_money(value)}</div>
      </div>"""


def _stat_tile(label, value, delta=None):
    delta_html = f'<div class="stat-delta">{_delta_span(delta)} vs prior</div>' if delta is not None else ""
    return f"""
      <div class="stat-tile">
        <div class="stat-label">{label}</div>
        <div class="stat-value">{value}</div>
        {delta_html}
      </div>"""


def _flag_card(flag):
    icon = "⛔" if flag["severity"] == "critical" else "⚠"
    return f"""
      <div class="flag-card flag-{flag['severity']}">
        <div class="flag-icon">{icon}</div>
        <div>
          <div class="flag-title">{flag['dimension'].title()}: {flag['name']}</div>
          <div class="flag-reason">{flag['reason']}</div>
        </div>
      </div>"""


def _df_to_table(df, columns, headers, formatters):
    thead = "".join(f"<th>{h}</th>" for h in headers)
    rows_html = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{formatters.get(c, str)(row[c])}</td>" for c in columns)
        rows_html.append(f"<tr>{cells}</tr>")
    return f"<table><thead><tr>{thead}</tr></thead><tbody>{''.join(rows_html)}</tbody></table>"


def write_html(path, summary_text, current_period, prior_period, generated_at,
                category_df, region_df, discount_product_df, discount_category_df, flags,
                total_revenue, total_profit, overall_margin, revenue_change):

    stat_tiles = "".join([
        _stat_tile("Total Revenue", _fmt_money(total_revenue), revenue_change),
        _stat_tile("Total Profit", _fmt_money(total_profit)),
        _stat_tile("Overall Margin", f"{overall_margin:.1f}%"),
        _stat_tile("Flags Raised", str(len(flags))),
    ])

    max_cat_rev = category_df["revenue"].max() if not category_df.empty else 0
    category_bars = "".join(_bar_row(r["category"], r["revenue"], max_cat_rev) for _, r in category_df.iterrows())

    max_region_rev = region_df["revenue"].max() if not region_df.empty else 0
    region_bars = "".join(_bar_row(r["region"], r["revenue"], max_region_rev) for _, r in region_df.iterrows())

    region_table = _df_to_table(
        region_df, ["region", "revenue", "prior_revenue", "pct_change", "margin"],
        ["Region", "Revenue", "Prior Revenue", "Change", "Margin"],
        {
            "revenue": _fmt_money, "prior_revenue": lambda v: _fmt_money(v) if pd.notna(v) else "n/a",
            "pct_change": _delta_span, "margin": lambda v: f"{v*100:.1f}%" if pd.notna(v) else "n/a",
        },
    )

    def discount_table(df):
        return _df_to_table(
            df, ["product" if "product" in df.columns else "category", "avg_discount", "revenue", "margin", "margin_risk"],
            ["Name", "Avg Discount", "Revenue", "Margin", "Risk"],
            {
                "avg_discount": lambda v: f"{v*100:.1f}%",
                "revenue": _fmt_money,
                "margin": lambda v: f"{v*100:.1f}%" if pd.notna(v) else "n/a",
                "margin_risk": lambda v: '<span class="risk-badge">⚠ margin risk</span>' if v else "—",
            },
        )

    flags_html = "".join(_flag_card(f) for f in flags) if flags else '<p class="muted">No flags raised this period.</p>'

    html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Sales Organizer Report - {current_period}</title>
<style>
  .viz-root {{
    color-scheme: light;
    --surface-1: #fcfcfb; --page: #f9f9f7; --text-primary: #0b0b0b;
    --text-secondary: #52514e; --text-muted: #898781; --gridline: #e1e0d9;
    --border: rgba(11,11,11,0.10); --series-1: #2a78d6;
    --delta-good: #006300; --delta-bad: #d03b3b;
    --status-warning: #fab219; --status-critical: #d03b3b;
  }}
  @media (prefers-color-scheme: dark) {{
    :root:where(:not([data-theme="light"])) .viz-root {{
      color-scheme: dark;
      --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff;
      --text-secondary: #c3c2b7; --text-muted: #898781; --gridline: #2c2c2a;
      --border: rgba(255,255,255,0.10); --series-1: #3987e5;
      --delta-good: #0ca30c; --delta-bad: #e66767;
    }}
  }}
  :root[data-theme="dark"] .viz-root {{
    color-scheme: dark;
    --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff;
    --text-secondary: #c3c2b7; --text-muted: #898781; --gridline: #2c2c2a;
    --border: rgba(255,255,255,0.10); --series-1: #3987e5;
    --delta-good: #0ca30c; --delta-bad: #e66767;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
          background: var(--page); color: var(--text-primary); }}
  .wrap {{ max-width: 980px; margin: 0 auto; padding: 32px 20px 64px; }}
  h1 {{ font-size: 1.6rem; margin: 0 0 4px; }}
  .meta {{ color: var(--text-secondary); font-size: 0.9rem; margin-bottom: 24px; }}
  .summary {{ background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px;
              padding: 18px 20px; line-height: 1.55; margin-bottom: 28px; }}
  h2 {{ font-size: 1.05rem; text-transform: uppercase; letter-spacing: 0.04em;
        color: var(--text-secondary); margin: 36px 0 12px; }}
  .stat-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; }}
  .stat-tile {{ background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }}
  .stat-label {{ color: var(--text-secondary); font-size: 0.8rem; text-transform: uppercase; letter-spacing: .03em; }}
  .stat-value {{ font-size: 1.6rem; font-weight: 600; margin-top: 4px; }}
  .stat-delta {{ margin-top: 6px; font-size: 0.85rem; }}
  .delta-up {{ color: var(--delta-good); font-weight: 600; }}
  .delta-down {{ color: var(--delta-bad); font-weight: 600; }}
  .muted {{ color: var(--text-muted); }}
  .bar-row {{ display: grid; grid-template-columns: 140px 1fr 90px; align-items: center; gap: 10px; margin: 8px 0; }}
  .bar-label {{ font-size: 0.88rem; color: var(--text-secondary); text-align: right; }}
  .bar-track {{ background: var(--gridline); border-radius: 4px; height: 14px; overflow: hidden; }}
  .bar-fill {{ background: var(--series-1); height: 100%; border-radius: 4px; }}
  .bar-value {{ font-size: 0.85rem; font-variant-numeric: tabular-nums; }}
  table {{ width: 100%; border-collapse: collapse; background: var(--surface-1);
           border: 1px solid var(--border); border-radius: 10px; overflow: hidden; font-size: 0.9rem; }}
  th, td {{ padding: 9px 12px; text-align: left; border-bottom: 1px solid var(--gridline); }}
  th {{ background: var(--text-primary); color: var(--surface-1); font-weight: 600; }}
  tr:last-child td {{ border-bottom: none; }}
  .risk-badge {{ color: var(--status-warning); font-weight: 600; }}
  .flag-card {{ display: flex; gap: 12px; align-items: flex-start; background: var(--surface-1);
                border: 1px solid var(--border); border-left: 4px solid var(--status-warning);
                border-radius: 8px; padding: 12px 14px; margin-bottom: 8px; }}
  .flag-critical {{ border-left-color: var(--status-critical); }}
  .flag-icon {{ font-size: 1.1rem; }}
  .flag-title {{ font-weight: 600; }}
  .flag-reason {{ color: var(--text-secondary); font-size: 0.88rem; margin-top: 2px; }}
  .cols-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 32px; }}
  @media (max-width: 720px) {{ .cols-2 {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>
<div class="viz-root">
<div class="wrap">
  <h1>Sales Organizer Report</h1>
  <div class="meta">Current period: <strong>{current_period}</strong>
    &nbsp;|&nbsp; Prior period: <strong>{prior_period if prior_period is not None else 'n/a'}</strong>
    &nbsp;|&nbsp; Generated {generated_at}</div>

  <div class="summary">{summary_text}</div>

  <h2>Overview</h2>
  <div class="stat-grid">{stat_tiles}</div>

  <div class="cols-2">
    <div>
      <h2>Revenue by Category</h2>
      {category_bars if category_bars else '<p class="muted">No data.</p>'}
    </div>
    <div>
      <h2>Revenue by Region</h2>
      {region_bars if region_bars else '<p class="muted">No data.</p>'}
    </div>
  </div>

  <h2>Region: Current vs. Prior Period</h2>
  {region_table}

  <h2>Biggest Discounts by Product</h2>
  {discount_table(discount_product_df)}

  <h2>Biggest Discounts by Category</h2>
  {discount_table(discount_category_df)}

  <h2>Flags</h2>
  {flags_html}
</div>
</div>
</body>
</html>"""

    with open(path, "w") as f:
        f.write(html)


# ---------------------------------------------------------------------------

def main():
    os.makedirs(REPORTS_DIR, exist_ok=True)

    data = common.load_data()
    current_df, prior_df, current_period, prior_period = common.split_periods(data)

    category_df = analysis.category_summary(current_df, prior_df)
    region_df = analysis.region_summary(current_df, prior_df)
    product_df = analysis.product_summary(current_df, prior_df)
    discount_product_df = analysis.discount_analysis(current_df, "product")
    discount_category_df = analysis.discount_analysis(current_df, "category")
    flags = analysis.generate_flags(category_df, region_df, product_df)
    flags_df = pd.DataFrame(flags) if flags else pd.DataFrame(columns=["dimension", "name", "reason", "severity"])

    summary_text = analysis.build_summary_paragraph(
        current_df, prior_df, current_period, prior_period, category_df, region_df, flags
    )

    total_revenue = current_df["revenue"].sum()
    total_profit = current_df["profit"].sum()
    overall_margin = (total_profit / total_revenue * 100) if total_revenue else 0
    prior_revenue = prior_df["revenue"].sum() if not prior_df.empty else None
    revenue_change = common.pct_change(total_revenue, prior_revenue) if prior_revenue else None

    stamp = str(current_period) if current_period is not None else datetime.now().strftime("%Y-%m-%d")
    xlsx_path = os.path.join(REPORTS_DIR, f"sales_report_{stamp}.xlsx")
    html_path = os.path.join(REPORTS_DIR, f"sales_report_{stamp}.html")

    write_excel(xlsx_path, summary_text, current_period, prior_period,
                category_df, region_df, discount_product_df, discount_category_df, flags_df)

    write_html(html_path, summary_text, current_period, prior_period,
               datetime.now().strftime("%Y-%m-%d %H:%M"),
               category_df, region_df, discount_product_df, discount_category_df, flags,
               total_revenue, total_profit, overall_margin, revenue_change)

    shutil.copyfile(xlsx_path, os.path.join(REPORTS_DIR, "latest.xlsx"))
    shutil.copyfile(html_path, os.path.join(REPORTS_DIR, "latest.html"))

    print(f"Report generated for {current_period} (prior: {prior_period}).")
    print(f"  Excel: {xlsx_path}")
    print(f"  HTML:  {html_path}")
    print(f"  Also updated: reports/latest.xlsx and reports/latest.html")


if __name__ == "__main__":
    main()
