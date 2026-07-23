"""Sales Organizer -- Streamlit web UI.

Run locally:
    streamlit run app.py

Lets you upload a weekly sales export (.xlsx) in the browser and see the
report immediately, instead of dropping the file into data/ and running
scripts/generate_report.py from the command line. Every step here --
validation, analysis, and the report itself -- calls straight into
common.py / analysis.py / generate_report.py, the same code the CLI uses,
so the two never drift apart.
"""

from __future__ import annotations

import io
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts"))

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import analysis
import common
import generate_report

st.set_page_config(page_title="Sales Organizer", page_icon="📊", layout="wide")

st.title("📊 Sales Organizer")

# --- Step 1: upload. Step 2: an explicit Generate Report click -- nothing
# else competes for attention above these two actions. Report state lives in
# session_state so it survives the reruns that later download-button clicks
# trigger, and resets if a different file is uploaded.
uploaded_file = st.file_uploader("Upload Data (.xlsx)", type=["xlsx"])

if uploaded_file is not None and st.session_state.get("uploaded_name") != uploaded_file.name:
    st.session_state["generated"] = False
    st.session_state["uploaded_name"] = uploaded_file.name

if uploaded_file is None:
    st.info(
        "Drop a `.xlsx` file above to get started. Expected columns: "
        "date, customer, product, category, region, quantity, price, discount, profit."
    )
    st.stop()

if st.button("Generate Report", type="primary", use_container_width=True):
    st.session_state["generated"] = True

if not st.session_state.get("generated"):
    st.stop()

# --- Run the file through the exact same pipeline the CLI uses ---
df, file_issues, halt_msg = common.process_file(uploaded_file, uploaded_file.name)

if halt_msg:
    st.error(f"🚫 {halt_msg}")
    st.stop()

data, issues = common.finalize_data([df], file_issues)
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

# --- Render the exact same report the CLI writes to reports/latest.html ---
# (no halts to show here -- a halted file already stopped above -- so pass [])
report_html = generate_report.render_html(
    summary_text, current_period, prior_period,
    datetime.now().strftime("%Y-%m-%d %H:%M"),
    category_df, region_df, discount_product_df, discount_category_df, flags,
    total_revenue, total_profit, overall_margin, revenue_change,
    issues, [],
)

if issues:
    st.warning(f"⚠ {len(issues)} data quality issue(s) were found -- see the banner in the report below.")

components.html(report_html, height=900, scrolling=True)

# --- Downloads, using the same write_excel()/render_html() the CLI uses ---
xlsx_buffer = io.BytesIO()
generate_report.write_excel(
    xlsx_buffer, summary_text, current_period, prior_period,
    category_df, region_df, discount_product_df, discount_category_df, flags_df,
    issues, [],
)

col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "⬇ Download Excel report", xlsx_buffer.getvalue(),
        file_name="sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col2:
    st.download_button(
        "⬇ Download HTML report", report_html,
        file_name="sales_report.html", mime="text/html",
    )
