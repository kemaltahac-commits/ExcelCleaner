import json

import pandas as pd

from excel_cleaner import clean_dataframe, process_file, unique_column_names


def sample_frame():
    return pd.DataFrame(
        {
            " customer  name ": ["  Ada   Lovelace ", "  Ada   Lovelace ", None, "Grace Hopper"],
            "EMAIL": [" ADA@EXAMPLE.COM ", " ADA@EXAMPLE.COM ", "bad-email", None],
            "Phone Number": ["(555) 123-4567", "(555) 123-4567", "12", None],
            "Order Date": ["2024-01-31", "2024-01-31", "03/04/2024", None],
            "Amount": ["1,200.50", "1,200.50", "30", None],
        }
    )


def test_clean_dataframe_cleans_expected_values():
    cleaned, report = clean_dataframe(sample_frame())
    assert list(cleaned.columns)[0] == "Customer Name"
    assert len(cleaned) == 3
    assert report["duplicates_removed"] == 1
    assert cleaned.iloc[0]["Customer Name"] == "Ada Lovelace"
    assert cleaned.iloc[0]["Email"] == "ada@example.com"
    assert cleaned.iloc[0]["Phone Number"] == "5551234567"
    assert cleaned.iloc[0]["Order Date"] == "2024-01-31"
    assert cleaned.iloc[1]["Order Date"] == "2024-03-04"
    assert report["invalid_emails"]["Email"] == ["bad-email"]
    assert report["invalid_phones"]["Phone Number"] == ["12"]
    assert "Order Date" not in report["invalid_dates"]
    assert report["numeric_conversions"]["Amount"] == 2
    assert report["missing_values_by_column"]["Email"] == 1


def test_empty_rows_are_removed():
    frame = pd.DataFrame({"Name": ["Ada", None], "Email": ["a@example.com", " "]})
    cleaned, report = clean_dataframe(frame)
    assert len(cleaned) == 1
    assert report["empty_rows_removed"] == 1


def test_similar_column_names_remain_separate():
    assert unique_column_names(["first name", "First_Name"]) == ["First Name", "First Name (2)"]


def test_us_numeric_and_currency_values_are_cleaned_without_touching_special_columns():
    frame = pd.DataFrame(
        {
            "Amount": ["1,250.50", "$1,250.50", "25,000", "2,000.00", "not numeric"],
            "Phone": ["555-123-4567", "555 123 4567", "+1 555 123 4567", "5551234567", "12"],
            "Email": ["a@example.com", "b@example.com", "c@example.com", "d@example.com", "e@example.com"],
            "Order Date": ["01/09/2026", "09/05/2026", "07/09/26", "2026-09-05", "13/45/2026"],
        }
    )
    cleaned, report = clean_dataframe(frame)
    assert cleaned["Amount"].tolist() == [1250.5, 1250.5, 25000.0, 2000.0, "not numeric"]
    assert report["numeric_conversions"]["Amount"] == 4
    assert report["invalid_numeric_values"]["Amount"] == ["not numeric"]
    assert cleaned["Phone"].tolist() == ["5551234567", "5551234567", "+15551234567", "5551234567", "12"]
    assert report["invalid_phones"]["Phone"] == ["12"]
    assert cleaned["Order Date"].tolist() == ["2026-01-09", "2026-09-05", "2026-07-09", "2026-09-05", "13/45/2026"]
    assert report["invalid_dates"]["Order Date"] == ["13/45/2026"]


def test_na_is_retained_but_empty_values_are_missing_and_near_duplicates_remain():
    frame = pd.DataFrame(
        {
            "Company": ["NA", "", "Acme", "Acme", "Acme Inc."],
            "Email": ["na@example.com", "empty@example.com", "a@example.com", "a@example.com", "a@example.com"],
        }
    )
    cleaned, report = clean_dataframe(frame)
    assert cleaned.iloc[0]["Company"] == "NA"
    assert report["missing_values_by_column"]["Company"] == 1
    assert report["duplicates_removed"] == 1
    assert len(cleaned) == 4
    assert "Acme Inc." in cleaned["Company"].tolist()


def test_process_file_creates_excel_and_report(tmp_path):
    source = tmp_path / "customers.xlsx"
    destination = tmp_path / "output" / "cleaned.xlsx"
    sample_frame().to_excel(source, index=False)
    output, report_path, report = process_file(source, destination)
    assert output.exists()
    assert report_path.exists()
    saved_report = json.loads(report_path.read_text(encoding="utf-8"))
    assert saved_report["output_file"] == str(destination)
    assert report["rows_after"] == 3


def test_process_file_preserves_literal_na_from_excel_input(tmp_path):
    source = tmp_path / "na_values.xlsx"
    pd.DataFrame({"Company": ["NA", ""], "Email": ["a@example.com", "b@example.com"]}).to_excel(source, index=False)
    _, _, report = process_file(source, tmp_path / "cleaned.xlsx")
    assert report["missing_values_by_column"]["Company"] == 1
