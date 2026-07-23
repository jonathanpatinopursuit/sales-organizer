#!/usr/bin/env python3
"""Generate the Sales Organizer report from every Excel file in data/.

Usage:
    ./run.sh
    (equivalent to: python3 scripts/generate_report.py)

Outputs (written to reports/):
    latest.xlsx / latest.html           - always overwritten; open these every time
    sales_report_<current-period>.xlsx  - dated copy, one sheet per section
    sales_report_<current-period>.html  - dated copy, single-file HTML report
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
                 category_df, region_df, discount_product_df, discount_category_df, flags_df,
                 issues=(), halts=()):
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
        halt_fmt = workbook.add_format({"bold": True, "font_color": "#ffffff", "bg_color": COLOR_CRITICAL})
        warn_fmt = workbook.add_format({"bold": True, "font_color": "#7a5b00", "bg_color": "#fef9c3"})
        skip_fmt = workbook.add_format({"bold": True, "font_color": COLOR_INK_SECONDARY, "bg_color": "#f3f4f6"})
        warn_text_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#fef9c3"})
        skip_text_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#f3f4f6"})
        halt_text_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#fee2e2"})

        # --- Summary sheet ---
        # Data Quality banner comes first (row 2 onward) -- before the summary
        # paragraph -- so a user sees whether the numbers below are trustworthy
        # before reading them, not after.
        ws = workbook.add_worksheet("Summary")
        writer.sheets["Summary"] = ws
        ws.set_column("A:A", 100)
        ws.write("A1", "Sales Organizer Report", workbook.add_format({"bold": True, "font_size": 16}))
        ws.write("A2", f"Current period: {current_period}   |   Prior period: {prior_period}")

        warn_issues = [i for i in issues if i["level"] == "warn"]
        skip_issues = [i for i in issues if i["level"] == "skip"]

        next_row = 2
        if halts:
            ws.write(next_row, 0, f"🚫 HALT — {len(halts)} file(s) rejected", halt_fmt)
            text = "\n".join(f"🚫 {h}" for h in halts)
            ws.write(next_row + 1, 0, text, halt_text_fmt)
            ws.set_row(next_row + 1, 16 * len(halts) + 10)
            next_row += 2
        if warn_issues:
            ws.write(next_row, 0, f"⚠ Warnings — {len(warn_issues)} item(s)", warn_fmt)
            text = "\n".join(f"⚠ {w['message']}" for w in warn_issues)
            ws.write(next_row + 1, 0, text, warn_text_fmt)
            ws.set_row(next_row + 1, 16 * len(warn_issues) + 10)
            next_row += 2
        if skip_issues:
            total_skipped = sum(s["count"] for s in skip_issues)
            ws.write(next_row, 0, f"⬜ Skipped — {total_skipped} row(s)", skip_fmt)
            text = "\n".join(f"⬜ {s['message']}" for s in skip_issues)
            ws.write(next_row + 1, 0, text, skip_text_fmt)
            ws.set_row(next_row + 1, 16 * len(skip_issues) + 10)
            next_row += 2

        ws.write(next_row, 0, summary_text, wrap_fmt)
        ws.set_row(next_row, 60)
        next_row += 1

        def write_table(df, sheet_name, pct_cols=(), money_cols=(), plain_pct_cols=(), flag_col=None):
            df = df.copy()
            # dq_note is an internal annotation, not a data column -- pull it out
            # before writing, keep it aligned by position for the inline-flag pass below.
            dq_notes = df.pop("dq_note").reset_index(drop=True) if "dq_note" in df.columns else None
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
            # inline flag: a row is only flagged if ITS OWN dq_note says a row that
            # actually contributed to that total was tagged -- not because of anything
            # else nearby (dq_note is computed per-group in analysis.py from the exact
            # rows behind that number).
            if flag_col and flag_col in df.columns and dq_notes is not None:
                col_idx = list(df.columns).index(flag_col)
                flag_cell_fmt = workbook.add_format({"font_color": COLOR_WARNING, "italic": True})
                for row_idx, note in enumerate(dq_notes):
                    if pd.notna(note):
                        val = df.iloc[row_idx][flag_col]
                        ws.write(2 + row_idx, col_idx, f"{val}  ⚠ {note}", flag_cell_fmt)
            return ws

        write_table(category_df, "By Category", pct_cols=["pct_change"],
                    money_cols=["revenue", "profit", "prior_revenue"], plain_pct_cols=["margin", "avg_discount"],
                    flag_col="category")
        write_table(region_df, "By Region", pct_cols=["pct_change"],
                    money_cols=["revenue", "profit", "prior_revenue"], plain_pct_cols=["margin", "avg_discount"],
                    flag_col="region")
        write_table(discount_product_df, "Discounts by Product", money_cols=["revenue", "profit"],
                    plain_pct_cols=["avg_discount", "margin"], flag_col="product")
        write_table(discount_category_df, "Discounts by Category", money_cols=["revenue", "profit"],
                    plain_pct_cols=["avg_discount", "margin"], flag_col="category")
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


def _name_with_note(value, dq_note):
    """Render a dimension name (product/category/region), with an inline flag
    only if dq_note is set -- dq_note is computed per-group in analysis.py from
    the exact rows behind that group's numbers, so this never flags a name just
    because something unrelated nearby had an issue."""
    if pd.notna(dq_note):
        return f'{value} <span class="dq-flag" title="{dq_note}">⚠️ <em>data adjusted</em></span>'
    return str(value)


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


def _flag_line(flag):
    """One plain-language sentence per flag (e.g. 'Sofa: margin only 6.7% and
    revenue down 18.2% vs prior period.') instead of a multi-part card, so
    Flags stays scannable at a glance rather than a dense table."""
    icon = "⛔" if flag["severity"] == "critical" else "⚠"
    reason = flag["reason"].replace("; ", " and ")
    return (
        f'<div class="flag-line flag-line-{flag["severity"]}">'
        f'{icon} <strong>{flag["dimension"].title()}: {flag["name"]}</strong> — {reason}.'
        f'</div>'
    )


def _df_to_table(df, columns, headers, formatters, name_col=None, note_col=None):
    thead = "".join(f"<th>{h}</th>" for h in headers)
    rows_html = []
    for _, row in df.iterrows():
        cells = []
        for c in columns:
            val = formatters.get(c, str)(row[c])
            if name_col and c == name_col and note_col:
                val = _name_with_note(row[c], row.get(note_col))
            cells.append(f"<td>{val}</td>")
        rows_html.append(f"<tr>{''.join(cells)}</tr>")
    return f"<table><thead><tr>{thead}</tr></thead><tbody>{''.join(rows_html)}</tbody></table>"


def render_dq_banner(issues=(), halts=()):
    """Build the 3-tier data-quality banner: HALT (red) / WARN (yellow) / SKIP (gray)."""
    if not issues and not halts:
        return ""

    warn_issues = [i for i in issues if i["level"] == "warn"]
    skip_issues = [i for i in issues if i["level"] == "skip"]

    rows = []
    for h in halts:
        rows.append(f'<div class="dq-halt">🚫 <strong>HALT:</strong> {h}</div>')
    if warn_issues:
        summary = "; ".join(w["message"] for w in warn_issues)
        rows.append(f'<div class="dq-warn">⚠️ <strong>{len(warn_issues)} warning(s)</strong> — {summary}</div>')
    if skip_issues:
        total_skipped = sum(s["count"] for s in skip_issues)
        summary = "; ".join(s["message"] for s in skip_issues)
        rows.append(f'<div class="dq-skip">⬜ <strong>{total_skipped} row(s) skipped</strong> — {summary}</div>')

    return f'<div id="dq-banner">{"".join(rows)}</div>'


def render_html(summary_text, current_period, prior_period, generated_at,
                 category_df, region_df, discount_product_df, discount_category_df, flags,
                 total_revenue, total_profit, overall_margin, revenue_change,
                 issues=(), halts=()) -> str:
    """Build the full HTML report and return it as a string. write_html()
    below is a thin wrapper that also saves it to disk -- this function is
    what app.py (the Streamlit UI) calls directly to embed the same report
    in the browser without writing a file, reusing 100% of this rendering
    logic rather than reimplementing it with different widgets."""

    data_quality_html = render_dq_banner(issues, halts)

    # Only the three headline numbers up top -- everything else (including a
    # flags count) is either shown as its own section below or cut, so the
    # first thing a user sees is three big numbers, not four competing tiles.
    stat_tiles = "".join([
        _stat_tile("Total Revenue", _fmt_money(total_revenue), revenue_change),
        _stat_tile("Total Profit", _fmt_money(total_profit)),
        _stat_tile("Overall Margin", f"{overall_margin:.1f}%"),
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
        name_col="region", note_col="dq_note",
    )

    def discount_table(df):
        name_col = "product" if "product" in df.columns else "category"

        return _df_to_table(
            df, [name_col, "avg_discount", "revenue", "margin", "margin_risk"],
            ["Name", "Avg Discount", "Revenue", "Margin", "Risk"],
            {
                "avg_discount": lambda v: f"{v*100:.1f}%",
                "revenue": _fmt_money,
                "margin": lambda v: f"{v*100:.1f}%" if pd.notna(v) else "n/a",
                "margin_risk": lambda v: '<span class="risk-badge">⚠ margin risk</span>' if v else "—",
            },
            name_col=name_col, note_col="dq_note",
        )

    flags_html = "".join(_flag_line(f) for f in flags) if flags else '<p class="muted">No flags raised this period.</p>'

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
    --dq-halt-bg: #fee2e2; --dq-halt-border: #dc2626; --dq-halt-text: #7f1d1d;
    --dq-warn-bg: #fef9c3; --dq-warn-border: #ca8a04; --dq-warn-text: #713f12;
    --dq-skip-bg: #f3f4f6; --dq-skip-border: #6b7280; --dq-skip-text: #374151;
  }}
  @media (prefers-color-scheme: dark) {{
    :root:where(:not([data-theme="light"])) .viz-root {{
      color-scheme: dark;
      --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff;
      --text-secondary: #c3c2b7; --text-muted: #898781; --gridline: #2c2c2a;
      --border: rgba(255,255,255,0.10); --series-1: #3987e5;
      --delta-good: #0ca30c; --delta-bad: #e66767;
      --dq-halt-bg: rgba(220,38,38,0.16); --dq-halt-border: #e66767; --dq-halt-text: #ffb4b4;
      --dq-warn-bg: rgba(202,138,4,0.18); --dq-warn-border: #fab219; --dq-warn-text: #ffdd8a;
      --dq-skip-bg: rgba(255,255,255,0.06); --dq-skip-border: #898781; --dq-skip-text: #c3c2b7;
    }}
  }}
  :root[data-theme="dark"] .viz-root {{
    color-scheme: dark;
    --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff;
    --text-secondary: #c3c2b7; --text-muted: #898781; --gridline: #2c2c2a;
    --border: rgba(255,255,255,0.10); --series-1: #3987e5;
    --delta-good: #0ca30c; --delta-bad: #e66767;
    --dq-halt-bg: rgba(220,38,38,0.16); --dq-halt-border: #e66767; --dq-halt-text: #ffb4b4;
    --dq-warn-bg: rgba(202,138,4,0.18); --dq-warn-border: #fab219; --dq-warn-text: #ffdd8a;
    --dq-skip-bg: rgba(255,255,255,0.06); --dq-skip-border: #898781; --dq-skip-text: #c3c2b7;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
          background: var(--page); color: var(--text-primary); }}
  .wrap {{ max-width: 980px; margin: 0 auto; padding: 32px 20px 64px; }}
  h1 {{ font-size: 1.6rem; margin: 0 0 4px; }}
  .meta {{ color: var(--text-secondary); font-size: 0.9rem; margin-bottom: 24px; }}
  .summary {{ background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px;
              padding: 18px 20px; line-height: 1.55; margin-bottom: 28px; }}
  #dq-banner {{ margin-bottom: 28px; }}
  .dq-halt, .dq-warn, .dq-skip {{ border-radius: 8px; padding: 10px 16px; margin-bottom: 6px; font-size: 0.88rem; }}
  .dq-halt {{ background: var(--dq-halt-bg); border-left: 4px solid var(--dq-halt-border); color: var(--dq-halt-text); }}
  .dq-warn {{ background: var(--dq-warn-bg); border-left: 4px solid var(--dq-warn-border); color: var(--dq-warn-text); }}
  .dq-skip {{ background: var(--dq-skip-bg); border-left: 4px solid var(--dq-skip-border); color: var(--dq-skip-text); }}
  .dq-flag {{ cursor: help; color: var(--status-warning); font-size: 0.85em; margin-left: 4px; }}
  h2 {{ font-size: 1.05rem; text-transform: uppercase; letter-spacing: 0.04em;
        color: var(--text-secondary); margin: 36px 0 12px; }}
  .stat-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; }}
  .stat-tile {{ background: var(--surface-1); border: 1px solid var(--border); border-radius: 12px; padding: 22px; }}
  .stat-label {{ color: var(--text-secondary); font-size: 0.85rem; text-transform: uppercase; letter-spacing: .03em; }}
  .stat-value {{ font-size: 2.4rem; font-weight: 700; margin-top: 6px; }}
  .stat-delta {{ margin-top: 8px; font-size: 0.9rem; }}
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
  .flag-line {{ background: var(--surface-1); border: 1px solid var(--border);
                border-left: 4px solid var(--status-warning); border-radius: 8px;
                padding: 10px 14px; margin-bottom: 6px; font-size: 0.92rem; }}
  .flag-line-critical {{ border-left-color: var(--status-critical); }}
  details.section {{ background: var(--surface-1); border: 1px solid var(--border);
                      border-radius: 10px; margin: 12px 0; overflow: hidden; }}
  details.section summary {{ cursor: pointer; padding: 14px 18px; font-weight: 600;
                              list-style: none; }}
  details.section summary::-webkit-details-marker {{ display: none; }}
  details.section summary::before {{ content: "▸  "; }}
  details.section[open] summary::before {{ content: "▾  "; }}
  details.section .section-body {{ padding: 4px 18px 18px; }}
</style>
</head>
<body>
<div class="viz-root">
<div class="wrap">
  <h1>Sales Organizer Report</h1>
  <div class="meta">Current period: <strong>{current_period}</strong>
    &nbsp;|&nbsp; Prior period: <strong>{prior_period if prior_period is not None else 'n/a'}</strong>
    &nbsp;|&nbsp; Generated {generated_at}</div>

  {data_quality_html}
  <div class="summary">{summary_text}</div>

  <div class="stat-grid">{stat_tiles}</div>

  <h2>Flags</h2>
  {flags_html}

  <details class="section">
    <summary>Revenue by Category</summary>
    <div class="section-body">
      {category_bars if category_bars else '<p class="muted">No data.</p>'}
    </div>
  </details>

  <details class="section">
    <summary>Revenue by Region</summary>
    <div class="section-body">
      {region_bars if region_bars else '<p class="muted">No data.</p>'}
      {region_table}
    </div>
  </details>

  <details class="section">
    <summary>Biggest Discounts by Product</summary>
    <div class="section-body">
      {discount_table(discount_product_df)}
    </div>
  </details>

  <details class="section">
    <summary>Biggest Discounts by Category</summary>
    <div class="section-body">
      {discount_table(discount_category_df)}
    </div>
  </details>
</div>
</div>
</body>
</html>"""

    return html


def write_html(path, summary_text, current_period, prior_period, generated_at,
                category_df, region_df, discount_product_df, discount_category_df, flags,
                total_revenue, total_profit, overall_margin, revenue_change,
                issues=(), halts=()):
    """Render the HTML report and save it to `path` (used by the CLI)."""
    html = render_html(
        summary_text, current_period, prior_period, generated_at,
        category_df, region_df, discount_product_df, discount_category_df, flags,
        total_revenue, total_profit, overall_margin, revenue_change,
        issues, halts,
    )
    with open(path, "w") as f:
        f.write(html)


# ---------------------------------------------------------------------------

def main():
    os.makedirs(REPORTS_DIR, exist_ok=True)

    data, issues, halts = common.load_data()
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
                category_df, region_df, discount_product_df, discount_category_df, flags_df,
                issues, halts)

    write_html(html_path, summary_text, current_period, prior_period,
               datetime.now().strftime("%Y-%m-%d %H:%M"),
               category_df, region_df, discount_product_df, discount_category_df, flags,
               total_revenue, total_profit, overall_margin, revenue_change,
               issues, halts)

    shutil.copyfile(xlsx_path, os.path.join(REPORTS_DIR, "latest.xlsx"))
    shutil.copyfile(html_path, os.path.join(REPORTS_DIR, "latest.html"))

    if halts:
        print(f"🚫 {len(halts)} file(s) could not be used:")
        for h in halts:
            print(f"   - {h}")
        print()

    if current_period is None:
        print("⚠ No usable data was found — fix the file(s) above and run this again.")
    else:
        period_note = f" (prior period: {prior_period})" if prior_period is not None else " (no prior period yet)"
        print(f"✅ Report generated for {current_period}{period_note}.")

    print(f"\n📄 Open this: {os.path.join(REPORTS_DIR, 'latest.html')}  (or latest.xlsx in Excel)")
    print(f"   Dated copy saved as: {os.path.basename(xlsx_path)}, {os.path.basename(html_path)}")

    if issues:
        print(f"\n⚠ {len(issues)} data quality issue(s) found — see the banner at the top of the report:")
        for i in issues:
            icon = "⚠" if i["level"] == "warn" else "⬜"
            print(f"   {icon} {i['message']}")


if __name__ == "__main__":
    main()
