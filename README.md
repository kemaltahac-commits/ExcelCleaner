# Excel Cleaner

Excel Cleaner is a command-line tool that safely cleans CSV and Excel (`.xlsx`) files and produces both a cleaned Excel file and a JSON change report. It never overwrites the source file and never invents missing data.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python excel_cleaner.py input/customers.xlsx
python excel_cleaner.py input/customers.csv --output output/customers_cleaned.xlsx
```

The default output for `input/customers.xlsx` is:

```text
output/customers_cleaned.xlsx
output/customers_cleaning_report.json
```

## What it cleans

- Readable, whitespace-normalized column names and string values
- Completely empty rows and exact duplicate rows
- Email casing and surrounding whitespace, while reporting invalid addresses
- Phone punctuation and spacing, while preserving digits and reporting suspicious values
- US-style dates to `YYYY-MM-DD`, including `MM/DD/YYYY`, `M/D/YY`, `YYYY-MM-DD`, and `YYYY/MM/DD`
- US numeric and currency formatting such as `1,200.50`, `25,000`, and `$1,250.50`
- Missing-value counts for every column

## Report

The JSON report includes row counts, removed duplicates and blank rows, missing-value counts, invalid email/phone/date values, standardization totals, numeric conversions or unconverted numeric-looking values, changed column names, and warnings.

## Limitations

The tool deliberately avoids guessing. Slash-separated dates use US month/day order by default. It does not infer missing values, country codes, or repairs for unknown email addresses. Literal `NA` is retained because it can be legitimate data. European numeric formats such as `1.250,50` are not supported. CSV files are read with standard comma-separated formatting.
