# Product Parser: Climate Shop

Python parser demo for product data enrichment from an online catalog.

## What it does

- Reads product model names from an Excel file.
- Opens search pages with Playwright.
- Finds matching product pages.
- Extracts product article/SKU data.
- Writes enriched results back to Excel.

## Stack

- Python
- asyncio
- Playwright
- openpyxl

## My role

Built the parsing flow, search navigation, product page extraction, retry-style checks and Excel result generation.

## Why this is relevant

This project is close to content/data roles: it combines website analysis, browser automation, product matching, Excel processing and data quality checks.

## Run

```bash
pip install -r requirements.txt
playwright install
python parser.py
```

The demo includes `sample_input.xlsx` as an example input file.

## What I would improve

- Add structured logging.
- Move settings to `.env`.
- Add CLI arguments for input/output files.
- Add stronger duplicate and empty-field checks.
- Add scheduled execution through cron or Windows Task Scheduler.
