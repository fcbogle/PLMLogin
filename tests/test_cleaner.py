"""Tests for login data cleaning."""

from __future__ import annotations

import pandas as pd

from config import AppConfig
from src.cleaner import clean_login_data, extract_user_display_name


def test_clean_login_data_parses_timestamps_and_derives_fields() -> None:
    raw_df = pd.DataFrame(
        {
            "User Name": ["Alice Smith (Alice Smith: Blatchford)", "Bob Jones (Bob Jones: Blatchford)"],
            "Event Time": ["2026-04-15 11:11:26 BST", "2026-01-03 08:05:00 GMT"],
        }
    )

    cleaned_df, report = clean_login_data(raw_df, AppConfig(excluded_user_exact_values=(), exclusions_file=None))

    assert len(cleaned_df) == 2
    assert report.input_rows == 2
    assert report.output_rows == 2
    assert report.dropped_invalid_timestamp_rows == 0
    assert cleaned_df["event_timestamp"].dt.strftime("%Y-%m-%d %H:%M:%S").tolist() == [
        "2026-04-15 11:11:26",
        "2026-01-03 08:05:00",
    ]
    assert cleaned_df["login_date"].dt.strftime("%Y-%m-%d").tolist() == ["2026-04-15", "2026-01-03"]
    assert cleaned_df["year_month"].tolist() == ["2026-04", "2026-01"]
    assert cleaned_df["user_display_name"].tolist() == ["Alice Smith", "Bob Jones"]


def test_clean_login_data_drops_missing_invalid_and_excluded_rows() -> None:
    raw_df = pd.DataFrame(
        {
            "User Name": [
                "Included User (Included User: Blatchford)",
                "WPS Test (WPS Test: Blatchford)",
                "Bad Timestamp (Bad Timestamp: Blatchford)",
                "   ",
            ],
            "Event Time": [
                "2026-04-15 11:11:26 BST",
                "2026-04-15 11:08:17 BST",
                "not-a-timestamp",
                "2026-04-15 10:58:55 BST",
            ],
        }
    )

    cleaned_df, report = clean_login_data(raw_df, AppConfig(exclusions_file=None))

    assert cleaned_df["user"].tolist() == ["Included User (Included User: Blatchford)"]
    assert report.input_rows == 4
    assert report.output_rows == 1
    assert report.dropped_missing_user_rows == 1
    assert report.dropped_excluded_user_rows == 1
    assert report.dropped_invalid_timestamp_rows == 1


def test_extract_user_display_name_handles_wrapped_and_plain_values() -> None:
    assert extract_user_display_name("Jane Doe (Jane Doe: Blatchford)") == "Jane Doe"
    assert extract_user_display_name("siteadmin") == "siteadmin"
