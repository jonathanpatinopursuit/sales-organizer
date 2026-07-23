"""Vercel entrypoint -- Flask wrapper around the existing report pipeline.

Vercel's Python runtime requires this exact file (api/app.py) to export a
top-level `app` variable. All the report logic already lives in scripts/
(common.py, analysis.py, generate_report.py) and is shared with the CLI
(scripts/generate_report.py) and the Streamlit UI (app.py) -- this file adds
no new logic, it just calls that same pipeline and serves the HTML string
generate_report.render_html() already returns.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts"))

from flask import Flask, request

import analysis
import common
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
  .viz-root {
    color-scheme: light;
    --surface-1: #fcfcfb; --page: #f9f9f7; --text-primary: #0b0b0b;
    --text-secondary: #52514e; --text-muted: #898781; --border: rgba(11,11,11,0.10);
    --series-1: #2a78d6; --status-critical: #d03b3b;
  }
  @media (prefers-color-scheme: dark) {
    :root:where(:not([data-theme="light"])) .viz-root {
      color-scheme: dark;
      --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff;
      --text-secondary: #c3c2b7; --text-muted: #898781; --border: rgba(255,255,255,0.10);
      --series-1: #3987e5; --status-critical: #e66767;
    }
  }
  :root[data-theme="dark"] .viz-root {
    color-scheme: dark;
    --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff;
    --text-secondary: #c3c2b7; --text-muted: #898781; --border: rgba(255,255,255,0.10);
    --series-1: #3987e5; --status-critical: #e66767;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
         background: var(--page); color: var(--text-primary); }
  .wrap { max-width: 640px; margin: 0 auto; padding: 60px 20px; }
  h1 { font-size: 1.8rem; margin: 0 0 20px; }
  .summary { background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px;
             padding: 18px 20px; line-height: 1.55; margin-bottom: 20px; }
  .upload-card { background: var(--surface-1); border: 1px solid var(--border); border-radius: 12px;
                 padding: 28px; }
  input[type=file] { display: block; width: 100%; margin-bottom: 18px; padding: 10px;
                      border: 1px solid var(--border); border-radius: 8px;
                      background: var(--page); color: var(--text-primary); }
  button.primary { width: 100%; padding: 14px; font-size: 1rem; font-weight: 600; border: none;
                    border-radius: 8px; background: var(--series-1); color: #fff; cursor: pointer; }
  button.primary:hover { opacity: 0.92; }
  .error { color: var(--status-critical); margin-bottom: 14px; font-size: 0.92rem; }
  .sample-link { display: block; text-align: center; margin-top: 16px; font-size: 0.88rem;
                  color: var(--text-secondary); }
"""


def render_upload_page(error: str | None = None) -> str:
    error_html = f'<div class="error">⚠ {error}</div>' if error else ""
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Sales Organizer</title>
<style>{UPLOAD_PAGE_CSS}</style>
</head>
<body>
<div class="viz-root"><div class="wrap">
  <h1>Sales Organizer</h1>
  {LANDING_HTML}
  <div class="upload-card">
    {error_html}
    <form method="POST" action="/generate" enctype="multipart/form-data">
      <input type="file" name="file" accept=".xlsx" required>
      <button type="submit" class="primary">Generate Report</button>
    </form>
  </div>
  <a class="sample-link" href="/sample">or view a sample report</a>
</div></div>
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


@app.route("/")
def index():
    return render_upload_page()


@app.route("/generate", methods=["POST"])
def generate():
    file = request.files.get("file")
    if file is None or file.filename == "":
        return render_upload_page(error="Choose a .xlsx file first.")

    try:
        df, file_issues, halt_msg = common.process_file(file, file.filename)
    except Exception as e:
        return render_upload_page(error=f"Couldn't read that file: {e}")

    if halt_msg:
        return render_upload_page(error=halt_msg)

    data, issues = common.finalize_data([df], file_issues)
    return _build_report_from_data(data, issues, [])


@app.route("/sample")
def sample():
    try:
        data, issues, halts = common.load_data()
    except FileNotFoundError as e:
        return f"<p>{e}</p>", 500
    return _build_report_from_data(data, issues, halts)
