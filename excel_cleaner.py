"""A small, conservative Excel/CSV data-cleaning command-line tool."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

import pandas as pd


EMAIL_COLUMN_PATTERN = re.compile(r"(^|[_\s])e-?mail($|[_\s])|email", re.IGNORECASE)
PHONE_COLUMN_PATTERN = re.compile(r"phone|mobile|telephone|tel|cell|gsm|fax", re.IGNORECASE)
DATE_COLUMN_PATTERN = re.compile(r"date|birth|dob|joined|created|updated|deadline", re.IGNORECASE)
EMAIL_PATTERN = re.compile(r"^[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?(?:\.[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?)+$", re.IGNORECASE)
# "NA" can be legitimate business data (for example, a company or region value).
MISSING_STRINGS = {"", "nan", "none", "null", "n/a"}


def normalize_whitespace(value: str) -> str:
    """Trim text and reduce runs of whitespace to one space."""
    return re.sub(r"\s+", " ", value.strip())


def normalize_column_name(name: Any) -> str:
    """Make headers readable without changing their meaning."""
    text = normalize_whitespace(str(name).replace("_", " ").replace("-", " "))
    return text.title() if text else "Unnamed Column"


def unique_column_names(columns: list[Any]) -> list[str]:
    """Normalize headings while preserving separate source columns."""
    used: dict[str, int] = {}
    result = []
    for column in columns:
        base = normalize_column_name(column)
        used[base] = used.get(base, 0) + 1
        result.append(base if used[base] == 1 else f"{base} ({used[base]})")
    return result


def is_missing(value: Any) -> bool:
    return pd.isna(value) or (isinstance(value, str) and value.strip().lower() in MISSING_STRINGS)


def clean_string_cells(frame: pd.DataFrame) -> pd.DataFrame:
    cleaned = frame.copy()
    for column in cleaned.columns:
        cleaned[column] = cleaned[column].map(
            lambda value: normalize_whitespace(value) if isinstance(value, str) else value
        )
    return cleaned


def is_email_column(column: str) -> bool:
    return bool(EMAIL_COLUMN_PATTERN.search(column))


def is_phone_column(column: str) -> bool:
    return bool(PHONE_COLUMN_PATTERN.search(column))


def is_date_column(column: str) -> bool:
    return bool(DATE_COLUMN_PATTERN.search(column))


def clean_email(value: Any) -> tuple[Any, bool]:
    if is_missing(value):
        return value, True
    if not isinstance(value, str):
        return value, False
    normalized = value.strip().lower()
    return normalized, bool(EMAIL_PATTERN.fullmatch(normalized))


def clean_phone(value: Any) -> tuple[Any, bool, bool]:
    """Return cleaned value, validity, and whether a safe standardization occurred."""
    if is_missing(value):
        return value, True, False
    if not isinstance(value, str):
        value = str(value)
    original = value
    compact = re.sub(r"[\s().-]+", "", value)
    if compact.startswith("00"):
        compact = "+" + compact[2:]
    elif compact.startswith("+"):
        compact = "+" + re.sub(r"\D", "", compact[1:])
    elif compact.isdigit():
        compact = compact
    else:
        return original, False, False
    digits = compact[1:] if compact.startswith("+") else compact
    valid = digits.isdigit() and 7 <= len(digits) <= 15
    return compact, valid, valid and compact != original


def looks_like_date_value(value: Any) -> bool:
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return True
    if not isinstance(value, str):
        return False
    return bool(re.search(r"[/-]", value) or re.fullmatch(r"\d{8}", value))


def parse_date_safely(value: Any) -> tuple[Any, bool, bool]:
    if is_missing(value):
        return value, True, False
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d"), True, True
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d"), True, True
    if not isinstance(value, str):
        return value, False, False
    text = value.strip()
    # Slash-separated dates deliberately use US month/day ordering.
    for date_format in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            parsed = datetime.strptime(text, date_format)
        except ValueError:
            continue
        normalized = parsed.strftime("%Y-%m-%d")
        return normalized, True, normalized != value
    return value, False, False


def parse_number_safely(value: Any) -> tuple[Any, bool]:
    if is_missing(value) or isinstance(value, bool) or isinstance(value, (int, float)):
        return value, False
    if not isinstance(value, str):
        return value, False
    text = value.strip()
    if not re.fullmatch(r"\$?[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", text):
        return value, False
    numeric = pd.to_numeric(text.replace("$", "").replace(",", ""), errors="coerce")
    if pd.isna(numeric):
        return value, False
    return numeric, True


def looks_numeric_column(series: pd.Series) -> bool:
    values = [value for value in series if not is_missing(value)]
    if not values:
        return False
    successes = sum(
        isinstance(value, (int, float)) and not isinstance(value, bool)
        or parse_number_safely(value)[1]
        for value in values
    )
    return successes / len(values) >= 0.8


def clean_dataframe(frame: pd.DataFrame, input_file: str = "") -> tuple[pd.DataFrame, dict[str, Any]]:
    """Clean a dataframe and return it with a client-readable change report."""
    if frame.empty and len(frame.columns) == 0:
        raise ValueError("The spreadsheet has no columns or rows.")

    rows_before = len(frame)
    original_columns = list(frame.columns)
    normalized_columns = unique_column_names(original_columns)
    frame = frame.copy()
    frame.columns = normalized_columns
    frame = clean_string_cells(frame)

    empty_mask = frame.apply(lambda row: all(is_missing(value) for value in row), axis=1)
    empty_rows_removed = int(empty_mask.sum())
    frame = frame.loc[~empty_mask].copy()
    before_duplicates = len(frame)
    frame = frame.drop_duplicates().copy()
    duplicates_removed = before_duplicates - len(frame)

    invalid_emails: dict[str, list[str]] = {}
    invalid_phones: dict[str, list[str]] = {}
    invalid_dates: dict[str, list[str]] = {}
    standardized_phone_count = 0
    standardized_date_count = 0
    numeric_conversions: dict[str, int] = {}
    invalid_numeric_values: dict[str, list[str]] = {}

    for column in frame.columns:
        if is_email_column(column):
            new_values = []
            invalid = []
            for value in frame[column]:
                cleaned, valid = clean_email(value)
                new_values.append(cleaned)
                if not valid and not is_missing(value):
                    invalid.append(str(value))
            frame[column] = new_values
            if invalid:
                invalid_emails[column] = invalid

        if is_phone_column(column):
            new_values = []
            invalid = []
            for value in frame[column]:
                cleaned, valid, standardized = clean_phone(value)
                new_values.append(cleaned)
                standardized_phone_count += int(standardized)
                if not valid and not is_missing(value):
                    invalid.append(str(value))
            frame[column] = new_values
            if invalid:
                invalid_phones[column] = invalid

        date_candidate = is_date_column(column) or sum(looks_like_date_value(value) for value in frame[column]) >= max(1, len(frame[column].dropna()) * 0.8)
        if date_candidate:
            new_values = []
            invalid = []
            for value in frame[column]:
                cleaned, valid, standardized = parse_date_safely(value)
                new_values.append(cleaned)
                standardized_date_count += int(standardized)
                if not valid and not is_missing(value):
                    invalid.append(str(value))
            frame[column] = new_values
            if invalid:
                invalid_dates[column] = invalid

        # Semantic values such as phone numbers must remain text, including leading zeroes.
        if not is_phone_column(column) and not is_email_column(column) and not date_candidate and looks_numeric_column(frame[column]):
            converted = 0
            values = []
            invalid = []
            for value in frame[column]:
                cleaned, was_converted = parse_number_safely(value)
                values.append(cleaned)
                converted += int(was_converted)
                if not was_converted and not is_missing(value) and not isinstance(value, (int, float)):
                    invalid.append(str(value))
            frame[column] = values
            if converted:
                numeric_conversions[column] = converted
            if invalid:
                invalid_numeric_values[column] = invalid

    missing_values = {column: int(sum(is_missing(value) for value in frame[column])) for column in frame.columns}
    column_changes = [
        {"original": str(old), "cleaned": new}
        for old, new in zip(original_columns, normalized_columns)
        if str(old) != new
    ]
    warnings = []
    if invalid_emails:
        warnings.append("Invalid email values were retained and listed in this report.")
    if invalid_phones:
        warnings.append("Suspicious phone values were retained and listed in this report.")
    if invalid_dates:
        warnings.append("Unrecognized or ambiguous dates were retained and listed in this report.")

    report = {
        "input_file": input_file,
        "output_file": "",
        "rows_before": rows_before,
        "rows_after": len(frame),
        "duplicates_removed": duplicates_removed,
        "empty_rows_removed": empty_rows_removed,
        "missing_values_by_column": missing_values,
        "invalid_emails": invalid_emails,
        "invalid_phones": invalid_phones,
        "invalid_dates": invalid_dates,
        "standardized_phone_count": standardized_phone_count,
        "standardized_date_count": standardized_date_count,
        "numeric_conversions": numeric_conversions,
        "invalid_numeric_values": invalid_numeric_values,
        "column_changes": column_changes,
        "warnings": warnings,
    }
    return frame, report


def read_input(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=object, keep_default_na=False)
    if path.suffix.lower() == ".xlsx":
        return pd.read_excel(path, dtype=object, engine="openpyxl", keep_default_na=False)
    raise ValueError("Unsupported file type. Please provide a .xlsx or .csv file.")


def default_output_path(input_path: Path) -> Path:
    return Path("output") / f"{input_path.stem}_cleaned.xlsx"


def process_file(input_path: str | Path, output_path: str | Path | None = None) -> tuple[Path, Path, dict[str, Any]]:
    source = Path(input_path)
    if not source.is_file():
        raise FileNotFoundError(f"Input file not found: {source}")
    destination = Path(output_path) if output_path else default_output_path(source)
    if destination.suffix.lower() != ".xlsx":
        raise ValueError("Output file must use the .xlsx extension.")
    destination.parent.mkdir(parents=True, exist_ok=True)
    frame = read_input(source)
    if frame.empty:
        raise ValueError("The spreadsheet contains no data rows.")
    cleaned, report = clean_dataframe(frame, str(source))
    report_stem = destination.stem[:-8] if destination.stem.endswith("_cleaned") else destination.stem
    report_path = destination.parent / f"{report_stem}_cleaning_report.json"
    report["output_file"] = str(destination)
    cleaned.to_excel(destination, index=False, engine="openpyxl")
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    return destination, report_path, report


def main() -> int:
    parser = argparse.ArgumentParser(description="Clean an Excel or CSV file without inventing data.")
    parser.add_argument("input_file", help="Path to the .xlsx or .csv file to clean")
    parser.add_argument("--output", help="Path for the cleaned .xlsx file")
    args = parser.parse_args()
    try:
        output_file, report_file, report = process_file(args.input_file, args.output)
    except (FileNotFoundError, PermissionError, ValueError, OSError, BadZipFile, UnicodeDecodeError, pd.errors.ParserError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    print(f"Cleaned {report['rows_before']} rows into {report['rows_after']} rows.")
    print(f"Created cleaned file: {output_file}")
    print(f"Created report: {report_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
