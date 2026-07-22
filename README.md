# Sales Organizer

Sales Organizer ingests your weekly sales order exports (Excel) and automatically
builds the sales summary a Sales Coordinator would otherwise put together by hand:
totals by category, region performance vs. the prior period, discount/margin risk,
underperformance flags, and a plain-language summary — plus a way to ask questions
about the data in plain English.

## What it does

Point it at one or more weekly Excel exports and it will:

- Total sales by product category
- Compare sales by region, current period vs. prior period (% change)
- Surface which products/categories are being discounted the most, and whether
  that discounting is hurting profit margin
- Automatically flag underperforming products, categories, or regions
- Write a short, auto-generated plain-language summary at the top of the report
- Let you ask questions like *"why is the West region down this month?"* and get
  an answer pulled straight from the data

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

## Adding new data

Every week, just drop your new export (`.xlsx`) into `data/`. You can keep old
exports there too — multiple files are combined automatically, and duplicate
weeks simply add up. Nothing in `data/` is tracked by git (see `.gitignore`),
so this is safe to do without worrying about committing spreadsheets.

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

## Asking questions

```bash
python3 scripts/ask.py "why is the West region down this month?"
python3 scripts/ask.py "how is the Electronics category doing?"
python3 scripts/ask.py "what's going on with Widget Pro?"
```

This looks for a region, category, or product name from your own data inside
the question, then pulls current-vs-prior-period numbers for that segment —
revenue change, margin change, the specific sub-segments (products/categories/
regions) driving the change, and any heavy-discount/margin-risk products.
It's rule-based and runs entirely on your local data — no API key required.

If it can't match anything specific in your question, it falls back to the
same general summary that leads the report, plus a list of known region/
category/product names you can ask about.

## Project structure

```
sales/
├── data/              your weekly Excel exports (gitignored, drop files here)
├── reports/            generated reports land here (gitignored, drop files here)
├── scripts/
│   ├── common.py        loads & normalizes data/*.xlsx
│   ├── analysis.py       category/region/discount/flag calculations (shared)
│   ├── generate_report.py   builds the Excel + HTML report
│   └── ask.py             the plain-language Q&A tool
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
4. **Smarter Q&A.** Wire `ask.py` up to an LLM (e.g. the Claude API) so it can
   answer free-form questions beyond named regions/categories/products —
   trends, forecasts, "what changed the most this week," etc. — while still
   grounding answers in the same underlying data.
5. **Data validation on intake.** Add a quick sanity check when a new export
   lands in `data/` (missing columns, duplicate rows, obviously bad dates) so
   bad data gets caught before it skews a report, instead of after.
