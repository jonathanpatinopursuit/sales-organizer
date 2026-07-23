# Sales Organizer

Sales Organizer ingests your weekly sales order exports (Excel) and automatically
builds the sales summary a Sales Coordinator would otherwise put together by hand:
totals by category, region performance vs. the prior period, discount/margin risk,
underperformance flags, and a plain-language summary.

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

| Problem | What happens |
|---|---|
| A required column is missing | **Halts** — that file is rejected with a clear error naming the missing column(s); other files in `data/` still process normally |
| More than 5% of rows have an unparseable/blank date | **Halts** — usually means the wrong column or a format the parser doesn't recognize; other files still process normally |
| A few rows (≤5%) have an unparseable/blank date | Those rows are **skipped**, rest of the file is used |
| A row has zero/negative quantity or a negative price | That row is **skipped** |
| A discount is outside 0–100% | **Clamped** to the nearest valid bound |
| Duplicate rows (within a file, or the same export saved under two filenames) | **Flagged only** — nothing is removed automatically, since legitimate repeat orders can look identical |

A halted file doesn't stop the run — it's excluded and the report still
generates from whatever files are valid (or, if every file halted, a report
still generates with $0 across the board and the banner explaining why).
Anything halted, skipped, or clamped shows up in a tiered "Data Quality"
banner at the top of the report (🚫 halt / ⚠ warning / ⬜ skipped, printed to
the console too), and rows whose discount was clamped or are duplicates get
an inline "⚠ data adjusted" flag next to that specific product/category in
the Discounts tables — not just an aggregate count — so problems are visible
instead of
silently changing your numbers.

## Adding new data

Every week, just drop your new export (`.xlsx`) into `data/`. You can keep old
exports there too — multiple files are combined automatically, and duplicate
weeks simply add up. Nothing in `data/` is tracked by git (see `.gitignore`),
so this is safe to do without worrying about committing spreadsheets.

Don't have a real export handy yet? Run `python3 scripts/create_sample_data.py`
to write a synthetic two-month `data/sample_sales.xlsx` (with a deliberately
weaker prior month for one region) so you can try the report and see the
period-over-period comparisons and flags in action.

## Setup

```bash
python3 -m pip install -r requirements.txt
```

## How to run

Generate the report (single command):

```bash
python3 scripts/generate_report.py
```

This reads every `.xlsx` file in `data/` and writes both an Excel workbook and
an HTML page to `reports/`:

- `reports/sales_report_<period>.xlsx` — multi-sheet workbook (Summary, By
  Category, By Region, Discounts by Product, Discounts by Category, Flags)
- `reports/sales_report_<period>.html` — single-file report you can open
  directly in a browser
- `reports/latest.xlsx` / `reports/latest.html` — always overwritten with the
  newest report, so you can bookmark one link/file

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
├── data/              your weekly Excel exports (gitignored, drop files here)
├── reports/            generated reports land here (gitignored, drop files here)
├── scripts/
│   ├── common.py        loads & normalizes data/*.xlsx
│   ├── analysis.py       category/region/discount/flag calculations (shared)
│   ├── generate_report.py   builds the Excel + HTML report
│   ├── validate_data.py     intake validation (bad dates, bad rows, discounts, dupes)
│   └── create_sample_data.py  writes a synthetic two-month sample export
├── requirements.txt
└── README.md
```

## Suggestions for extending this further

1. **Schedule it.** Add a weekly cron job / macOS `launchd` job (or a GitHub
   Actions workflow) that runs `generate_report.py` automatically once the
   week's export lands, so the report is always waiting for you Monday morning.
2. **Alerts.** Have the script post to Slack or send an email automatically
   whenever a new `critical` flag appears (e.g. a region down >30%, or a
   category running negative profit).
3. **A live dashboard.** Swap the static HTML report for a small always-on
   dashboard (e.g. Streamlit or a simple web app) so you can filter by date
   range, region, or category interactively instead of regenerating files.
