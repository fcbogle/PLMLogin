"""ADU denial audit loading, cleaning, and analysis."""

from __future__ import annotations

import logging

import pandas as pd

from config import AppConfig
from src.cleaner import build_excluded_user_mask, extract_user_display_name

LOGGER = logging.getLogger(__name__)


def load_adu_denials(config: AppConfig) -> pd.DataFrame:
    """Load the optional ADU denial audit workbook."""

    if config.adu_input_file is None:
        return build_empty_adu_raw_data()

    if not config.adu_input_file.exists():
        raise FileNotFoundError(f"ADU input file not found: {config.adu_input_file}")

    LOGGER.info(
        "Loading ADU denial workbook %s [sheet=%s]",
        config.adu_input_file,
        config.adu_input_sheet,
    )
    dataframe = pd.read_excel(config.adu_input_file, sheet_name=config.adu_input_sheet)
    dataframe.columns = [str(column).strip() for column in dataframe.columns]
    return dataframe


def clean_adu_denials(raw_df: pd.DataFrame, config: AppConfig) -> pd.DataFrame:
    """Validate and clean ADU denial audit rows."""

    if raw_df.empty:
        return build_empty_adu_raw_data()

    required_columns = {
        config.adu_user_column,
        config.adu_timestamp_column,
        config.adu_event_label_column,
    }
    missing_columns = required_columns - set(raw_df.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Missing required ADU columns: {missing}")

    df = raw_df.copy()
    df.rename(
        columns={
            config.adu_user_column: "user",
            config.adu_timestamp_column: "event_time_raw",
            config.adu_event_label_column: "event_label",
        },
        inplace=True,
    )
    df = df.loc[df["event_label"] == config.adu_denied_event_label].copy()
    df["user"] = df["user"].astype("string").str.strip()
    missing_user_mask = df["user"].isna() | (df["user"] == "")
    df = df.loc[~missing_user_mask].copy()

    excluded_user_mask = build_excluded_user_mask(df["user"], config)
    df = df.loc[~excluded_user_mask].copy()
    df["user_display_name"] = df["user"].apply(extract_user_display_name)

    df["event_time_text"] = df["event_time_raw"].astype("string").str.strip()
    for suffix in config.timestamp_suffixes_to_strip:
        df["event_time_text"] = df["event_time_text"].str.removesuffix(suffix)

    df["event_timestamp"] = pd.to_datetime(
        df["event_time_text"],
        format="%Y-%m-%d %H:%M:%S",
        errors="coerce",
    )
    df = df.loc[df["event_timestamp"].notna()].copy()
    df["denial_date"] = df["event_timestamp"].dt.normalize()
    df["year"] = df["event_timestamp"].dt.year
    df["month"] = df["event_timestamp"].dt.month
    df["year_month"] = df["event_timestamp"].dt.to_period("M").astype(str)
    return df.sort_values(["event_timestamp", "user"]).reset_index(drop=True)


def build_adu_user_summary(cleaned_adu_df: pd.DataFrame, config: AppConfig) -> pd.DataFrame:
    """Build per-user ADU denial metrics."""

    if cleaned_adu_df.empty:
        return pd.DataFrame(
            columns=[
                "user",
                "user_display_name",
                "adu_denied_attempts",
                "adu_denied_days",
                "adu_denied_months",
                "first_adu_denial",
                "last_adu_denial",
                "adu_denied_attempts_last_90_days",
                "adu_denied_days_last_90_days",
                "is_repeated_adu_denial_user",
            ]
        )

    max_timestamp = cleaned_adu_df["event_timestamp"].max()
    recent_cutoff = max_timestamp.normalize() - pd.Timedelta(days=config.recent_activity_days)
    recent_df = cleaned_adu_df.loc[cleaned_adu_df["denial_date"] >= recent_cutoff]

    summary = cleaned_adu_df.groupby("user").agg(
        user_display_name=("user_display_name", "first"),
        adu_denied_attempts=("event_timestamp", "size"),
        adu_denied_days=("denial_date", "nunique"),
        adu_denied_months=("year_month", "nunique"),
        first_adu_denial=("denial_date", "min"),
        last_adu_denial=("denial_date", "max"),
    )
    recent_attempts = recent_df.groupby("user")["event_timestamp"].size().rename("adu_denied_attempts_last_90_days")
    recent_days = recent_df.groupby("user")["denial_date"].nunique().rename("adu_denied_days_last_90_days")
    summary = summary.join(recent_attempts).join(recent_days)
    summary[["adu_denied_attempts_last_90_days", "adu_denied_days_last_90_days"]] = summary[
        ["adu_denied_attempts_last_90_days", "adu_denied_days_last_90_days"]
    ].fillna(0).astype(int)
    summary["is_repeated_adu_denial_user"] = (
        (summary["adu_denied_days"] >= config.adu_repeated_denial_days_threshold)
        | (summary["adu_denied_attempts"] >= config.adu_repeated_denial_attempts_threshold)
    )
    return summary.reset_index().sort_values(
        ["adu_denied_days", "adu_denied_attempts", "last_adu_denial", "user_display_name"],
        ascending=[False, False, False, True],
    ).reset_index(drop=True)


def build_adu_monthly_denials(cleaned_adu_df: pd.DataFrame) -> pd.DataFrame:
    """Build ADU denial trend metrics by month."""

    if cleaned_adu_df.empty:
        return pd.DataFrame(
            columns=["month_label", "year_month", "adu_denied_attempts", "adu_denied_users", "adu_denied_days"]
        )

    monthly = cleaned_adu_df.groupby("year_month").agg(
        adu_denied_attempts=("event_timestamp", "size"),
        adu_denied_users=("user", "nunique"),
        adu_denied_days=("denial_date", "nunique"),
    ).reset_index().sort_values("year_month").reset_index(drop=True)
    monthly.insert(0, "month_label", [f"M{index}" for index in range(1, len(monthly) + 1)])
    return monthly


def build_empty_adu_raw_data() -> pd.DataFrame:
    """Return an empty cleaned ADU dataframe with expected columns."""

    return pd.DataFrame(
        columns=[
            "event_label",
            "event_time_raw",
            "user",
            "user_display_name",
            "event_timestamp",
            "denial_date",
            "year",
            "month",
            "year_month",
        ]
    )
