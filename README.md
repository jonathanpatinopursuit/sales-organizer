# Sales Organizer

Sales Organizer ingests your weekly sales order exports (Excel) and automatically
builds the sales summary a Sales Coordinator would otherwise put together by hand:
totals by category, region performance vs. the prior period, discount/margin risk,
underperformance flags, and a plain-language summary.

## Quick start

```bash
python3 -m pip install -r requirements.txt   # first time only
./run.sh
```

1. Drop your weekly `.xlsx` export into `data/`.
2. Run `./run.sh`.
3. Open **`reports/latest.html`** (or `reports/latest.xlsx` in Excel) — always
   the newest report, same filename every time.

If anything's wrong with your file, the terminal and the top of the report
will say exactly what's wrong and how to fix it.

Prefer a browser to the command line? See [Web UI (Streamlit)](#web-ui-streamlit) below.

**Live report:** https://jonathanpatinopursuit.github.io/sales/
(publishing this is a manual step — see [Publishing the live report](#publishing-the-live-report) —
so it only ever shows a report you've explicitly chosen to make public.)

## What it does

Point it at one or more weekly Excel exports and it will:

- Total sales by product category
- Compare sales by region, current period vs. prior period (% change)
- Surface which products/categories are being discounted the most, and whether
  that discounting is hurting profit margin
- Automatically flag underperforming products, categories, or regions
- Write a short, auto-generated plain-language summary at the top of the report

"Current period" and "prior period" are the two most recent calendar months
found in your data (based on the `date` column) — so as you drop in new weekly
exports, the report automatically rolls forward.

## Expected input file format

Drop `.xlsx` files into `data/`. Each file needs these columns (case-insensitive,
any column order):

| column     | description                                             |
|------------|----------------------------------------------------------|
| `date`     | order date                                              |
| `customer` | customer name                                           |
| `product`  | product name                                            |
| `category` | product category                                        |
| `region`   | sales region                                            |
| `quantity` | units sold                                              |
| `price`    | unit price                                              |
| `discount` | discount rate applied (e.g. `0.1` for 10%, or `10`)     |
| `profit`   | total profit for that line item                         |

Revenue is derived as `quantity * price * (1 - discount)`. You don't need to
compute revenue yourself.

## Data quality checks

Every file is validated on intake (`scripts/validate_data.py`) before it's used:

| Problem | dq_flag label | What happens |
|---|---|---|
| A required column is missing | — | **Halts** — that file is rejected with a clear error naming the missing column(s); other files in `data/` still process normally |
| More than 5% of rows have an unparseable/blank date | — | **Halts** — usually means the wrong column or a format the parser doesn't recognize; other files still process normally |
| A few rows (≤5%) have an unparseable/blank date | `skipped:bad_date` | Those rows are **skipped**, rest of the file is used |
| A row has zero/negative quantity or a negative price | `skipped:invalid_qty_price` | That row is **skipped** |
| A row has a blank/missing region | `flagged:invalid_region` | **Kept** (dropping it would lose real revenue), just flagged |
| A row has a blank/missing category | `flagged:invalid_category` | **Kept**, just flagged |
| A discount is outside 0–100% | `clamped:discount` | **Clamped** to the nearest valid bound, row kept |
| A row has negative profit | `flagged:negative_profit` | **Kept** — not necessarily wrong (could be a real loss), but worth surfacing |
| Duplicate rows (within a file, or the same export saved under two filenames) | `flagged:duplicate` | **Kept** — nothing is removed automatically, since legitimate repeat orders can look identical |

Every row gets tagged with its own `dq_flag` (skipped rows are tagged before
being dropped, purely so the check that dropped them can be tested in
isolation — a dropped row never reaches any metric). A row hit by more than
one check keeps every label. A halted file doesn't stop the run — it's
excluded and the report still generates from whatever files are valid (or,
if every file halted, a report still generates with $0 across the board and
the banner explaining why).

Anything halted, skipped, or clamped shows up in a tiered "Data Quality"
banner at the top of the report (🚫 halt / ⚠ warning / ⬜ skipped, printed to
the console too). Separately, every category/region/product total in the
report carries its own inline flag computed from the *exact rows that went
into that number* (see `analysis.py`'s `dq_note`, built from `dq_flag`
grouped the same way as the metric itself) — so a flag always means "this
number was computed using at least one tagged row," never "something nearby
had an issue." A product with an issue only in the *prior* period leaves its
*current*-period total unflagged, since that number was never computed from
the affected row — and a category is flagged only when the group actually
contains a tagged row, even if a different product in that same category is
perfectly clean.

Run `python3 scripts/test_validation.py` to exercise every check in
isolation (halts, skips, clamps, flags, and a row hit by two checks at
once) — it prints a ✅/❌ per assertion and exits non-zero on failure.

## Adding new data

Every week, just drop your new export (`.xlsx`) into `data/`. You can keep old
exports there too — multiple files are combined automatically, and duplicate
weeks simply add up. Nothing in `data/` is tracked by git (see `.gitignore`),
so this is safe to do without worrying about committing spreadsheets.

Don't have a real export handy yet? Run `python3 scripts/create_sample_data.py`
to write a synthetic two-month `data/sample_sales.xlsx` (with a deliberately
weaker prior month for one region) so you can try the report and see the
period-over-period comparisons and flags in action.

## Output files

Every run of `./run.sh` writes both an Excel workbook and an HTML page to
`reports/`:

- **`reports/latest.xlsx`** / **`reports/latest.html`** — always overwritten
  with the newest report. This is what you open every time; the name never
  changes, so you never have to search for it.
- `reports/sales_report_<period>.xlsx` / `.html` — a dated copy of the same
  report (e.g. `sales_report_2026-07.html`), kept alongside `latest.*` in
  case you want to look back at an older period.

The Excel workbook has one sheet per section (Summary, By Category, By
Region, Discounts by Product, Discounts by Category, Flags); the HTML report
is a single self-contained file.

## Web UI (Streamlit)

`app.py` is a browser-based alternative to the command line: upload a file
instead of dropping it into `data/`, and the report renders on the page
immediately (plus download buttons for the Excel/HTML files). It runs the
exact same `common.py` / `analysis.py` / `generate_report.py` code as the
CLI, so the report is identical either way — nothing is reimplemented.

Run locally:

```bash
python3 -m pip install -r requirements.txt   # first time only
streamlit run app.py
```

This opens the app at `http://localhost:8501`.

**Deploying to Streamlit Community Cloud** (this is a manual step you do
once via Streamlit's own site, the same way [publishing the live
report](#publishing-the-live-report) is a manual step — nothing here
deploys automatically):

1. Push this repo to GitHub (already done — it's this repo).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in, and
   click **New app**.
3. Point it at this repo, branch `main`, main file path `app.py`.
4. Streamlit Cloud installs everything from `requirements.txt`
   automatically — no extra configuration needed.

One difference from the CLI: the web UI processes one uploaded file at a
time. The CLI can combine multiple weekly exports at once from `data/` (see
[Adding new data](#adding-new-data)); upload them one at a time here, or use
`./run.sh` locally if you need to combine several.

## Publishing the live report

This repo is **public**, so `reports/` is deliberately gitignored — nothing
generated from your real sales data is ever pushed automatically. When you
*do* want the live link (above) to reflect your latest report, publish it
explicitly:

```bash
./scripts/publish_report.sh
```

This copies `reports/latest.html` to `docs/index.html`, commits, and pushes —
GitHub Pages serves whatever is in `docs/` at
https://jonathanpatinopursuit.github.io/sales/. Only run this when
`reports/latest.html` contains data you're OK with being publicly visible.

## Project structure

```
sales/
├── run.sh              generates the report -- the only command you need
├── app.py              Streamlit web UI (upload a file, see the report in a browser)
├── data/              your weekly Excel exports (gitignored, drop files here)
├── reports/            generated reports land here (gitignored, drop files here)
├── scripts/
│   ├── common.py        loads & normalizes data/*.xlsx
│   ├── analysis.py       category/region/discount/flag calculations (shared)
│   ├── generate_report.py   builds the Excel + HTML report
│   ├── validate_data.py     intake validation (bad dates, bad rows, discounts, dupes)
│   ├── test_validation.py   test suite for validate_data.py
│   └── create_sample_data.py  writes a synthetic two-month sample export
├── requirements.txt   (also read by Streamlit Community Cloud at deploy time)
└── README.md
```

## Suggestions for extending this further

1. **Schedule it.** Add a weekly cron job / macOS `launchd` job (or a GitHub
   Actions workflow) that runs `generate_report.py` automatically once the
   week's export lands, so the report is always waiting for you Monday morning.
2. **Alerts.** Have the script post to Slack or send an email automatically
   whenever a new `critical` flag appears (e.g. a region down >30%, or a
   category running negative profit).
3. **Interactive filtering in the web UI.** `app.py` currently renders the
   same static report the CLI produces; adding filter widgets (date range,
   region, category) that recompute `analysis.py`'s tables live would turn
   it into more of a dashboard instead of a one-shot upload-and-view.
