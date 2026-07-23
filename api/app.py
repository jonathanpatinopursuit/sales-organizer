"""Vercel entrypoint -- Flask wrapper around the existing report pipeline.

Vercel's Python runtime requires this exact file (api/app.py) to export a
top-level `app` variable. All the report logic already lives in scripts/
(common.py, analysis.py, generate_report.py) and is shared with the CLI
(scripts/generate_report.py) and the Streamlit UI (app.py) -- this file adds
no new logic, it just calls that same pipeline and serves the HTML string
generate_report.render_html() already returns.
"""

from __future__ import annotations

import io
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts"))

from flask import Flask, request

import analysis
import common
import create_sample_data
import generate_report

app = Flask(__name__)

LANDING_HTML = (
    '<div class="summary" style="margin-bottom:20px;">'
    '<strong>Sales Organizer</strong> turns a weekly sales export (.xlsx) into an instant report: '
    'revenue and margin by category and region, biggest discounts by product, and automatic flags '
    'for declining segments or thin margins.'
    '</div>'
)

UPLOAD_PAGE_CSS = """
  /* A till-receipt aesthetic, on purpose: a sales report *is* the thing a
  receipt promises. Committed to one look (paper-light, monospace) rather
  than adapting to light/dark -- a printed receipt doesn't have a dark mode. */
  * { box-sizing: border-box; }
  html, body { margin: 0; }
  body {
    background: #d8d8d2;
    padding: 40px 16px;
    display: flex;
    justify-content: center;
    font-family: ui-monospace, "SF Mono", "Menlo", "Consolas", "Liberation Mono", monospace;
  }
  .receipt {
    background: #fbfbf7;
    color: #1c1c1a;
    width: 100%;
    max-width: 380px;
    padding: 30px 26px 26px;
    -webkit-mask-image:
      radial-gradient(circle 6px at 0 0, transparent 6px, black 6.5px),
      radial-gradient(circle 6px at 100% 0, transparent 6px, black 6.5px);
    -webkit-mask-position: top left, top right;
    -webkit-mask-size: 51% 100%;
    -webkit-mask-repeat: no-repeat;
  }
  .receipt .center { text-align: center; }
  .receipt .brand { font-size: 1.15rem; font-weight: 700; letter-spacing: 0.18em; margin: 0 0 4px; }
  .receipt .sub { font-size: 0.72rem; letter-spacing: 0.14em; color: #8d8d86; margin: 0 0 18px; }
  .receipt .divider { border: none; border-top: 1px dashed #b9b9b2; margin: 16px 0; }
  .receipt p.lede { font-size: 0.8rem; line-height: 1.6; margin: 0 0 18px; }
  .receipt .error-line { color: #a23c3c; font-size: 0.8rem; font-weight: 700; margin: 0 0 14px; }
  .receipt .line-row { display: flex; align-items: baseline; gap: 6px; font-size: 0.82rem; margin: 14px 0; }
  .receipt .line-row .label { white-space: nowrap; letter-spacing: 0.03em; }
  .receipt .line-row .leader { flex: 1; border-bottom: 1px dotted #b9b9b2; transform: translateY(-4px); }
  .receipt input[type="file"] { font-family: inherit; font-size: 0.72rem; max-width: 150px; color: #1c1c1a; }
  .receipt .total-rule { border: none; border-top: 2px solid #1c1c1a; border-bottom: 2px solid #1c1c1a; height: 4px; margin: 4px 0 0; }
  .receipt button.total {
    display: block; width: 100%; margin-top: 20px; padding: 12px;
    font-family: inherit; font-weight: 700; font-size: 0.86rem; letter-spacing: 0.06em;
    background: #1c1c1a; color: #fbfbf7; border: none; cursor: pointer;
  }
  .receipt button.total:hover { background: #3a3a36; }
  .receipt .sample-row { display: flex; justify-content: space-between; align-items: center; font-size: 0.76rem; margin-top: 8px; }
  .receipt button.sample-link {
    font-family: inherit; font-size: 0.76rem; background: none; border: none;
    color: #a23c3c; text-decoration: underline; cursor: pointer; padding: 0;
  }
  .receipt .barcode {
    height: 34px; margin-top: 22px;
    background: repeating-linear-gradient(
      90deg, #1c1c1a 0 2px, transparent 2px 3px, #1c1c1a 3px 4px,
      transparent 4px 7px, #1c1c1a 7px 9px, transparent 9px 11px
    );
  }
  .receipt .footer-note { text-align: center; font-size: 0.66rem; letter-spacing: 0.08em; color: #8d8d86; margin-top: 10px; }
"""


def render_upload_page(error: str | None = None) -> str:
    error_html = f'<p class="error-line">⚠ {error}</p>' if error else ""
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Sales Organizer</title>
<style>{UPLOAD_PAGE_CSS}</style>
</head>
<body>
<div class="receipt">
  <div class="center">
    <p class="brand">SALES ORGANIZER</p>
    <p class="sub">— WEEKLY REPORT —</p>
  </div>
  <hr class="divider">
  {error_html}
  <p class="lede">Revenue and margin by category and region, biggest discounts, and flags for anything declining -- from one uploaded export.</p>
  <form method="POST" action="/generate" enctype="multipart/form-data">
    <div class="line-row">
      <span class="label">UPLOAD FILE</span>
      <span class="leader"></span>
      <input type="file" name="file" accept=".xlsx" required>
    </div>
    <hr class="total-rule">
    <button class="total" type="submit">GENERATE REPORT</button>
  </form>
  <hr class="divider">
  <form method="POST" action="/sample" style="display:contents;">
    <div class="sample-row">
      <span>NO DATA ON HAND?</span>
      <button class="sample-link" type="submit">generate a sample report</button>
    </div>
  </form>
  <div class="barcode"></div>
  <p class="footer-note">KEEP THIS REPORT FOR YOUR RECORDS</p>
</div>
</body>
</html>"""


def _build_report_from_data(data, issues, halts) -> str:
    current_df, prior_df, current_period, prior_period = common.split_periods(data)

    category_df = analysis.category_summary(current_df, prior_df)
    region_df = analysis.region_summary(current_df, prior_df)
    product_df = analysis.product_summary(current_df, prior_df)
    discount_product_df = analysis.discount_analysis(current_df, "product")
    discount_category_df = analysis.discount_analysis(current_df, "category")
    flags = analysis.generate_flags(category_df, region_df, product_df)

    summary_text = analysis.build_summary_paragraph(
        current_df, prior_df, current_period, prior_period, category_df, region_df, flags
    )

    total_revenue = current_df["revenue"].sum()
    total_profit = current_df["profit"].sum()
    overall_margin = (total_profit / total_revenue * 100) if total_revenue else 0
    prior_revenue = prior_df["revenue"].sum() if not prior_df.empty else None
    revenue_change = common.pct_change(total_revenue, prior_revenue) if prior_revenue else None

    html = generate_report.render_html(
        summary_text, current_period, prior_period,
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        category_df, region_df, discount_product_df, discount_category_df, flags,
        total_revenue, total_profit, overall_margin, revenue_change,
        issues, halts,
    )

    # Slot the intro card and a way back to the upload form in above the
    # report, reusing the report's own ".summary" styling so it matches
    # without adding any new CSS.
    anchor = '<div class="wrap">\n  <h1>Sales Organizer Report</h1>'
    back_link = '<p><a href="/">← Upload a different file</a></p>'
    if anchor in html:
        html = html.replace(
            anchor,
            f'<div class="wrap">\n  {LANDING_HTML}\n  {back_link}\n  <h1>Sales Organizer Report</h1>',
            1,
        )
    return html


def _generate_report_from_file(file_like, filename: str) -> str:
    """Shared by /generate (a real upload) and /sample (synthetic data) so
    both run through the exact same read/validate/analyze pipeline -- the
    only difference is where the raw .xlsx bytes come from."""
    try:
        df, file_issues, halt_msg = common.process_file(file_like, filename)
    except Exception as e:
        return render_upload_page(error=f"Couldn't read that file: {e}")

    if halt_msg:
        return render_upload_page(error=halt_msg)

    data, issues = common.finalize_data([df], file_issues)
    return _build_report_from_data(data, issues, [])


@app.route("/")
def index():
    return render_upload_page()


@app.route("/generate", methods=["POST"])
def generate():
    file = request.files.get("file")
    if file is None or file.filename == "":
        return render_upload_page(error="Choose a .xlsx file first.")
    return _generate_report_from_file(file, file.filename)


@app.route("/sample", methods=["GET", "POST"])
def sample():
    """No raw data on hand? Build a synthetic two-period dataset in memory
    and run it through the same pipeline as a real upload -- no file needed."""
    buf = io.BytesIO()
    create_sample_data.sample_dataframe().to_excel(buf, index=False)
    buf.seek(0)
    return _generate_report_from_file(buf, "sample_sales.xlsx")
