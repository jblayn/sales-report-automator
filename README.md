# Sales Report Automator

Turns a messy sales export (CSV or Excel) into a clean, formatted Excel report with summary tables and charts in one command. It replaces hours of manual copy-paste and pivot tables every week.

## What it does

- **Cleans the data:** fixes inconsistent capitalization and stray spaces, reads mixed date formats (`2026-03-01`, `03/01/2026`, `01-Mar-2026`), strips `$` and commas from prices, and removes duplicate rows.
- **Flags bad rows:** rows with missing or invalid values go to a separate *Rejected Rows* sheet so nothing disappears without a trace.
- **Builds the report:** a KPI summary plus revenue broken down by month, region, product, and sales rep, each with a chart.

## Demo

```
$ python report.py sample_data/sales_raw.csv -o sample_data/sales_report.xlsx
Report saved to sample_data/sales_report.xlsx
  510 rows in -> 476 clean rows (10 duplicates, 24 invalid removed)
  Total revenue: $268,844.42
```

The finished report is in `sample_data/sales_report.xlsx`.

| Sheet | Contents |
|---|---|
| Summary | Total revenue, orders, units, average order value, top performers, cleaning stats |
| By Month | Monthly revenue and a line chart |
| By Region / By Product / By Sales Rep | Ranked tables and bar charts |
| Clean Data | The full cleaned dataset with a calculated Revenue column |
| Rejected Rows | Rows that couldn't be used, for review |

## Setup

```
pip install -r requirements.txt
python make_sample_data.py      # optional: regenerate the messy demo data
python report.py <your_file.csv> -o report.xlsx
```

## Expected columns

`Order ID, Order Date, Region, Sales Rep, Product, Quantity, Unit Price`. Column names are matched loosely, so capitalization and spacing don't matter. For a client's data, you only need to change the column mapping in `clean()`.

## Tech

Python, pandas, openpyxl
