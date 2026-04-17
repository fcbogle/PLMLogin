"""Analysis functions for PLM login records."""

from __future__ import annotations

from dataclasses import asdict

import pandas as pd

from config import AppConfig
from src.categoriser import categorise_user
from src.models import AnalysisOutputs, CleaningReport


def build_analysis_outputs(
    cleaned_df: pd.DataFrame,
    config: AppConfig,
    cleaning_report: CleaningReport,
) -> AnalysisOutputs:
    """Build the full set of workbook outputs from cleaned login data."""

    user_summary = build_user_summary(cleaned_df, config)
    monthly_activity = build_monthly_activity(cleaned_df)
    category_summary = build_category_summary(user_summary)
    overview_metrics = build_overview_metrics(cleaned_df, user_summary, category_summary, cleaning_report)
    category_rules = build_category_rules(config)

    return AnalysisOutputs(
        raw_data=cleaned_df.sort_values(["event_timestamp", "user"]).reset_index(drop=True),
        user_summary=user_summary,
        regular_users=user_summary.loc[user_summary["usage_category"] == "Regular"].reset_index(drop=True),
        occasional_users=user_summary.loc[user_summary["usage_category"] == "Occasional"].reset_index(drop=True),
        rare_users=user_summary.loc[user_summary["usage_category"] == "Rare"].reset_index(drop=True),
        monthly_activity=monthly_activity,
        overview_metrics=overview_metrics,
        category_summary=category_summary,
        category_rules=category_rules,
        cleaning_report=cleaning_report,
    )


def build_user_summary(cleaned_df: pd.DataFrame, config: AppConfig) -> pd.DataFrame:
    """Build per-user summary metrics."""

    reporting_months = cleaned_df["year_month"].nunique()
    max_timestamp = cleaned_df["event_timestamp"].max()

    user_daily_activity = (
        cleaned_df[["user", "login_date"]]
        .drop_duplicates()
        .sort_values(["user", "login_date"])
        .reset_index(drop=True)
    )
    user_daily_activity["previous_login_date"] = user_daily_activity.groupby("user")["login_date"].shift()
    user_daily_activity["gap_days"] = (
        user_daily_activity["login_date"] - user_daily_activity["previous_login_date"]
    ).dt.days

    user_login_counts = cleaned_df.groupby("user").agg(
        user_display_name=("user_display_name", "first"),
        total_logins=("event_timestamp", "size"),
        active_months=("year_month", "nunique"),
        first_login_date=("login_date", "min"),
        last_login_date=("login_date", "max"),
    )

    distinct_login_days = (
        cleaned_df.groupby("user")["login_date"].nunique().rename("distinct_login_days")
    )

    gap_stats = user_daily_activity.groupby("user").agg(
        longest_gap_days=("gap_days", "max"),
        median_days_between_activity=("gap_days", "median"),
    )

    summary = user_login_counts.join(distinct_login_days).join(gap_stats)
    summary["average_logins_per_month"] = summary["total_logins"] / reporting_months
    summary["average_active_days_per_month"] = summary["distinct_login_days"] / reporting_months
    summary["percentage_months_active"] = (summary["active_months"] / reporting_months * 100).round(2)
    recent_cutoff = max_timestamp.normalize() - pd.Timedelta(days=config.recent_activity_days)
    summary["last_90_days_active"] = summary["last_login_date"] >= recent_cutoff
    summary["longest_gap_days"] = summary["longest_gap_days"].fillna(0).astype(int)
    summary["median_days_between_activity"] = summary["median_days_between_activity"].fillna(0).round(2)
    summary["usage_category"] = summary["average_active_days_per_month"].apply(
        lambda value: categorise_user(value, config.thresholds)
    )

    summary = summary.reset_index()
    preferred_order = [
        "user",
        "user_display_name",
        "total_logins",
        "distinct_login_days",
        "active_months",
        "average_logins_per_month",
        "average_active_days_per_month",
        "percentage_months_active",
        "first_login_date",
        "last_login_date",
        "longest_gap_days",
        "median_days_between_activity",
        "last_90_days_active",
        "usage_category",
    ]
    summary = summary[preferred_order].sort_values(
        ["usage_category", "average_active_days_per_month", "distinct_login_days", "user_display_name", "user"],
        ascending=[True, False, False, True, True],
    )
    return summary.reset_index(drop=True)


def build_monthly_activity(cleaned_df: pd.DataFrame) -> pd.DataFrame:
    """Build a user-by-month matrix of distinct active days."""

    monthly = (
        cleaned_df.groupby(["user", "user_display_name", "year_month"])["login_date"]
        .nunique()
        .reset_index(name="active_days")
    )
    matrix = monthly.pivot(
        index=["user", "user_display_name"],
        columns="year_month",
        values="active_days",
    ).fillna(0).astype(int)
    matrix["Total"] = matrix.sum(axis=1)
    matrix = matrix.reset_index()
    return matrix.sort_values(["user_display_name", "user"]).reset_index(drop=True)


def build_category_summary(user_summary: pd.DataFrame) -> pd.DataFrame:
    """Build category counts and percentages for the overview sheet."""

    total_users = len(user_summary)
    summary = (
        user_summary.groupby("usage_category")
        .size()
        .rename("user_count")
        .reindex(["Regular", "Occasional", "Rare"], fill_value=0)
        .reset_index()
        .rename(columns={"usage_category": "category"})
    )
    summary["percentage_of_total_users"] = (summary["user_count"] / total_users * 100).round(2)
    return summary


def build_overview_metrics(
    cleaned_df: pd.DataFrame,
    user_summary: pd.DataFrame,
    category_summary: pd.DataFrame,
    cleaning_report: CleaningReport,
) -> pd.DataFrame:
    """Build headline overview metrics."""

    category_counts = dict(zip(category_summary["category"], category_summary["user_count"], strict=False))
    metrics = [
        ("total_login_records", len(cleaned_df)),
        ("total_unique_users", user_summary["user"].nunique()),
        ("date_range_start", cleaned_df["login_date"].min().date().isoformat()),
        ("date_range_end", cleaned_df["login_date"].max().date().isoformat()),
        ("regular_user_count", category_counts.get("Regular", 0)),
        ("occasional_user_count", category_counts.get("Occasional", 0)),
        ("rare_user_count", category_counts.get("Rare", 0)),
        ("average_logins_per_user", round(user_summary["total_logins"].mean(), 2)),
        ("average_active_days_per_user", round(user_summary["distinct_login_days"].mean(), 2)),
    ]

    for field_name, value in asdict(cleaning_report).items():
        metrics.append((field_name, value))

    return pd.DataFrame(metrics, columns=["metric", "value"])


def build_category_rules(config: AppConfig) -> pd.DataFrame:
    """Document category thresholds used by the run."""

    rules = [
        {
            "category": "Regular",
            "rule_description": "average_active_days_per_month >= regular threshold",
            "threshold_value": config.thresholds.regular_min_average_active_days,
        },
        {
            "category": "Occasional",
            "rule_description": "average_active_days_per_month >= occasional threshold and below regular threshold",
            "threshold_value": config.thresholds.occasional_min_average_active_days,
        },
        {
            "category": "Rare",
            "rule_description": "average_active_days_per_month < occasional threshold",
            "threshold_value": config.thresholds.occasional_min_average_active_days,
        },
    ]
    return pd.DataFrame(rules)
