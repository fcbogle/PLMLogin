"""Data cleaning for PLM login records."""

from __future__ import annotations

import logging
import re

import pandas as pd

from config import AppConfig
from src.models import CleaningReport

LOGGER = logging.getLogger(__name__)


def clean_login_data(raw_df: pd.DataFrame, config: AppConfig) -> tuple[pd.DataFrame, CleaningReport]:
    """Validate and clean the raw login export."""

    df = raw_df.copy()
    required_columns = {config.user_column, config.timestamp_column}
    missing_columns = required_columns - set(df.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Missing required columns: {missing}")

    df.rename(
        columns={
            config.user_column: "user",
            config.timestamp_column: "event_time_raw",
        },
        inplace=True,
    )

    df["user"] = df["user"].astype("string").str.strip()
    if config.normalise_user_case:
        df["user"] = df["user"].str.lower()

    missing_user_mask = df["user"].isna() | (df["user"] == "")
    dropped_missing_user_rows = int(missing_user_mask.sum())
    df = df.loc[~missing_user_mask].copy()
    df["user_display_name"] = df["user"].apply(extract_user_display_name)

    excluded_user_mask = build_excluded_user_mask(df["user"], config)
    dropped_excluded_user_rows = int(excluded_user_mask.sum())
    df = df.loc[~excluded_user_mask].copy()

    df["event_time_text"] = df["event_time_raw"].astype("string").str.strip()
    for suffix in config.timestamp_suffixes_to_strip:
        df["event_time_text"] = df["event_time_text"].str.removesuffix(suffix)

    df["event_timestamp"] = pd.to_datetime(
        df["event_time_text"],
        format="%Y-%m-%d %H:%M:%S",
        errors="coerce",
    )

    invalid_timestamp_mask = df["event_timestamp"].isna()
    dropped_invalid_timestamp_rows = int(invalid_timestamp_mask.sum())
    df = df.loc[~invalid_timestamp_mask].copy()

    df["login_date"] = df["event_timestamp"].dt.normalize()
    df["year"] = df["event_timestamp"].dt.year
    df["month"] = df["event_timestamp"].dt.month
    df["year_month"] = df["event_timestamp"].dt.to_period("M").astype(str)

    report = CleaningReport(
        input_rows=len(raw_df),
        output_rows=len(df),
        dropped_missing_user_rows=dropped_missing_user_rows,
        dropped_invalid_timestamp_rows=dropped_invalid_timestamp_rows,
        dropped_excluded_user_rows=dropped_excluded_user_rows,
    )

    LOGGER.info(
        "Cleaned rows: input=%s output=%s dropped_missing_user=%s dropped_invalid_timestamp=%s dropped_excluded_user=%s",
        report.input_rows,
        report.output_rows,
        report.dropped_missing_user_rows,
        report.dropped_invalid_timestamp_rows,
        report.dropped_excluded_user_rows,
    )
    return df, report


def extract_user_display_name(user_value: str) -> str:
    """Extract a cleaner display name from the raw export field."""

    match = re.match(r"^(.*?)\s+\(", user_value)
    if match:
        return match.group(1).strip()
    return user_value.strip()


def build_excluded_user_mask(user_series: pd.Series, config: AppConfig) -> pd.Series:
    """Return a boolean mask for rows excluded from analysis."""

    user_casefold = user_series.astype("string").str.casefold()
    exact_values = {value.casefold() for value in config.excluded_user_exact_values}
    contains_values = tuple(value.casefold() for value in config.excluded_user_contains_values)

    exact_mask = user_casefold.isin(exact_values) if exact_values else pd.Series(False, index=user_series.index)
    if contains_values:
        contains_mask = user_casefold.apply(lambda value: any(token in value for token in contains_values))
    else:
        contains_mask = pd.Series(False, index=user_series.index)
    return exact_mask | contains_mask
