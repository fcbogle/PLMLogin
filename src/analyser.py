"""Analysis functions for PLM login records."""

from __future__ import annotations

from dataclasses import asdict

import pandas as pd

from config import AppConfig
from src.categoriser import categorise_user
from src.models import AnalysisOutputs, CleaningReport
from src.production_technicians import (
    build_production_technician_match_report,
    normalise_person_name,
)


def build_analysis_outputs(
    cleaned_df: pd.DataFrame,
    config: AppConfig,
    cleaning_report: CleaningReport,
    production_technicians: pd.DataFrame | None = None,
    cleaned_adu_df: pd.DataFrame | None = None,
    adu_user_summary: pd.DataFrame | None = None,
    adu_monthly_denials: pd.DataFrame | None = None,
) -> AnalysisOutputs:
    """Build the full set of workbook outputs from cleaned login data."""

    user_summary = build_user_summary(cleaned_df, config)
    user_summary = add_production_technician_flags(user_summary, production_technicians)
    adu_user_summary = adu_user_summary if adu_user_summary is not None else pd.DataFrame()
    cleaned_adu_df = cleaned_adu_df if cleaned_adu_df is not None else pd.DataFrame()
    adu_monthly_denials = adu_monthly_denials if adu_monthly_denials is not None else pd.DataFrame()
    user_summary = add_adu_flags(user_summary, adu_user_summary, config)
    monthly_activity = build_monthly_activity(cleaned_df)
    category_summary = build_category_summary(user_summary)
    monthly_active_users = build_monthly_active_users(cleaned_df)
    most_active_users = build_most_active_users(user_summary)
    at_risk_rare_users = build_at_risk_rare_users(user_summary)
    production_technician_matches = build_production_technician_match_report(
        production_technicians if production_technicians is not None else pd.DataFrame(),
        user_summary,
    )
    production_technician_review_summary = build_production_technician_review_summary(
        production_technician_matches
    )
    licence_recommendations = build_licence_recommendations(
        user_summary,
        adu_user_summary,
        production_technician_matches,
        config,
    )
    dedicated_licence_candidates = build_dedicated_licence_candidates(licence_recommendations)
    recommendation_summary = build_recommendation_summary(licence_recommendations)
    licence_balance = build_licence_balance(licence_recommendations)
    unused_licence_evidence = build_unused_licence_evidence(licence_recommendations)
    executive_summary = build_executive_summary(
        cleaned_df,
        user_summary,
        adu_user_summary,
        production_technician_matches,
        licence_recommendations,
    )
    overview_metrics = build_overview_metrics(cleaned_df, user_summary, category_summary, cleaning_report)
    category_rules = build_category_rules(config)

    return AnalysisOutputs(
        raw_data=cleaned_df.sort_values(["event_timestamp", "user"]).reset_index(drop=True),
        user_summary=user_summary,
        regular_users=user_summary.loc[user_summary["usage_category"] == "Regular"].reset_index(drop=True),
        occasional_users=user_summary.loc[user_summary["usage_category"] == "Occasional"].reset_index(drop=True),
        rare_users=user_summary.loc[user_summary["usage_category"] == "Rare"].reset_index(drop=True),
        adu_raw_data=cleaned_adu_df.sort_values(["event_timestamp", "user"]).reset_index(drop=True)
        if not cleaned_adu_df.empty
        else cleaned_adu_df,
        adu_user_summary=adu_user_summary,
        adu_monthly_denials=adu_monthly_denials,
        monthly_activity=monthly_activity,
        licence_recommendations=licence_recommendations,
        dedicated_licence_candidates=dedicated_licence_candidates,
        unused_licence_evidence=unused_licence_evidence,
        executive_summary=executive_summary,
        recommendation_summary=recommendation_summary,
        production_technician_review_summary=production_technician_review_summary,
        licence_balance=licence_balance,
        overview_metrics=overview_metrics,
        category_summary=category_summary,
        monthly_active_users=monthly_active_users,
        most_active_users=most_active_users,
        at_risk_rare_users=at_risk_rare_users,
        production_technician_matches=production_technician_matches,
        category_rules=category_rules,
        cleaning_report=cleaning_report,
    )


def add_adu_flags(
    user_summary: pd.DataFrame,
    adu_user_summary: pd.DataFrame,
    config: AppConfig,
) -> pd.DataFrame:
    """Add ADU denial metrics to the login user summary."""

    summary = user_summary.copy()
    adu_columns = [
        "adu_denied_attempts",
        "adu_denied_days",
        "adu_denied_months",
        "last_adu_denial",
        "adu_denied_attempts_last_90_days",
        "adu_denied_days_last_90_days",
        "is_repeated_adu_denial_user",
    ]
    for column in adu_columns:
        summary[column] = 0 if column != "last_adu_denial" else pd.NaT

    if adu_user_summary.empty:
        summary["has_adu_denials"] = False
        summary["is_dedicated_licence_candidate"] = False
        return summary

    adu = adu_user_summary.copy()
    adu["match_key"] = adu["user_display_name"].apply(normalise_person_name)
    summary["match_key"] = summary["user_display_name"].apply(normalise_person_name)
    summary = summary.merge(
        adu[
            [
                "match_key",
                "adu_denied_attempts",
                "adu_denied_days",
                "adu_denied_months",
                "last_adu_denial",
                "adu_denied_attempts_last_90_days",
                "adu_denied_days_last_90_days",
                "is_repeated_adu_denial_user",
            ]
        ],
        on="match_key",
        how="left",
        suffixes=("", "_adu"),
    )
    for column in adu_columns:
        merged_column = f"{column}_adu"
        if merged_column in summary.columns:
            summary[column] = summary[merged_column].combine_first(summary[column])
            summary.drop(columns=[merged_column], inplace=True)

    numeric_columns = [
        "adu_denied_attempts",
        "adu_denied_days",
        "adu_denied_months",
        "adu_denied_attempts_last_90_days",
        "adu_denied_days_last_90_days",
    ]
    summary[numeric_columns] = summary[numeric_columns].fillna(0).astype(int)
    summary["is_repeated_adu_denial_user"] = summary["is_repeated_adu_denial_user"].fillna(False).astype(bool)
    summary["has_adu_denials"] = summary["adu_denied_attempts"] > 0
    summary["is_dedicated_licence_candidate"] = (
        (summary["adu_denied_days"] >= config.dedicated_licence_denied_days_threshold)
        & summary["usage_category"].isin(["Regular", "Occasional"])
    )
    return summary.drop(columns=["match_key"])


def add_production_technician_flags(
    user_summary: pd.DataFrame,
    production_technicians: pd.DataFrame | None,
) -> pd.DataFrame:
    """Flag users who match the optional Production Technician list."""

    summary = user_summary.copy()
    summary["is_production_technician"] = False
    summary["production_technician_name"] = pd.NA

    if production_technicians is None or production_technicians.empty:
        return summary

    technician_lookup = production_technicians.set_index("production_technician_match_key")[
        "production_technician_name"
    ].to_dict()
    summary_match_key = summary["user_display_name"].apply(normalise_person_name)
    summary["production_technician_name"] = summary_match_key.map(technician_lookup)
    summary["is_production_technician"] = summary["production_technician_name"].notna()
    return summary


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


def build_monthly_active_users(cleaned_df: pd.DataFrame) -> pd.DataFrame:
    """Build monthly active user counts for reporting charts."""

    monthly_active_users = (
        cleaned_df.groupby("year_month")["user"]
        .nunique()
        .reset_index(name="active_users")
        .sort_values("year_month")
        .reset_index(drop=True)
    )
    monthly_active_users.insert(
        0,
        "month_label",
        [f"M{index}" for index in range(1, len(monthly_active_users) + 1)],
    )
    return monthly_active_users


def build_at_risk_rare_users(user_summary: pd.DataFrame, limit: int = 50) -> pd.DataFrame:
    """Build a table of rare users who have not been active in the last 90 days."""

    at_risk = user_summary.loc[
        (user_summary["usage_category"] == "Rare") & (~user_summary["last_90_days_active"])
    ].sort_values(
        [
            "average_active_days_per_month",
            "distinct_login_days",
            "total_logins",
            "last_login_date",
            "user_display_name",
            "user",
        ],
        ascending=[True, True, True, True, True, True],
    ).head(limit)

    return at_risk[
        [
            "user_display_name",
            "usage_category",
            "average_active_days_per_month",
            "distinct_login_days",
            "total_logins",
            "last_login_date",
            "last_90_days_active",
        ]
    ].reset_index(drop=True)


def build_production_technician_review_summary(production_technician_matches: pd.DataFrame) -> pd.DataFrame:
    """Summarise Production Technician usage review statuses."""

    if production_technician_matches.empty:
        return pd.DataFrame(columns=["review_status", "user_count"])

    review = production_technician_matches.copy()
    review["review_status"] = review.apply(classify_production_technician_review_status, axis=1)
    return review.groupby("review_status").size().rename("user_count").reset_index().sort_values(
        "review_status"
    ).reset_index(drop=True)


def classify_production_technician_review_status(row: pd.Series) -> str:
    """Classify a Production Technician for licence review."""

    if row["match_status"] != "Matched":
        return "No observed PLM authentication"
    if row["usage_category"] == "Regular":
        return "Retain dedicated licence"
    if row["usage_category"] == "Occasional":
        return "Review business need"
    return "Candidate for licence removal/reallocation"


def build_licence_recommendations(
    user_summary: pd.DataFrame,
    adu_user_summary: pd.DataFrame,
    production_technician_matches: pd.DataFrame,
    config: AppConfig,
) -> pd.DataFrame:
    """Build the combined licence recommendation table."""

    rows: list[dict[str, object]] = []
    user_lookup = user_summary.copy()
    user_lookup["match_key"] = user_lookup["user_display_name"].apply(normalise_person_name)
    user_lookup = user_lookup.set_index("match_key", drop=False)

    adu_lookup = adu_user_summary.copy()
    if not adu_lookup.empty:
        adu_lookup["match_key"] = adu_lookup["user_display_name"].apply(normalise_person_name)
        adu_lookup = adu_lookup.set_index("match_key", drop=False)

    relevant_keys = set()
    if not adu_lookup.empty:
        relevant_keys.update(adu_lookup.index.tolist())
    if not production_technician_matches.empty:
        relevant_keys.update(
            production_technician_matches["production_technician_name"].apply(normalise_person_name).tolist()
        )

    for match_key in sorted(relevant_keys):
        login_row = user_lookup.loc[match_key] if match_key in user_lookup.index else None
        if isinstance(login_row, pd.DataFrame):
            login_row = login_row.iloc[0]
        adu_row = adu_lookup.loc[match_key] if not adu_lookup.empty and match_key in adu_lookup.index else None
        if isinstance(adu_row, pd.DataFrame):
            adu_row = adu_row.iloc[0]

        technician_rows = production_technician_matches.loc[
            production_technician_matches["production_technician_name"].apply(normalise_person_name) == match_key
        ] if not production_technician_matches.empty else pd.DataFrame()
        technician_row = technician_rows.iloc[0] if not technician_rows.empty else None

        user_name = choose_recommendation_user_name(login_row, adu_row, technician_row)
        is_production_technician = technician_row is not None
        has_adu_denials = adu_row is not None
        usage_category = get_optional_value(login_row, "usage_category")
        adu_denied_attempts = int(get_optional_value(adu_row, "adu_denied_attempts", 0) or 0)
        adu_denied_days = int(get_optional_value(adu_row, "adu_denied_days", 0) or 0)
        recommendation, rationale = classify_licence_recommendation(
            usage_category=usage_category,
            is_production_technician=is_production_technician,
            production_match_status=get_optional_value(technician_row, "match_status"),
            has_adu_denials=has_adu_denials,
            adu_denied_attempts=adu_denied_attempts,
            adu_denied_days=adu_denied_days,
            config=config,
        )
        rows.append(
            {
                "user": user_name,
                "user_group": build_user_group_label(is_production_technician, has_adu_denials),
                "usage_category": usage_category or "No observed PLM authentication",
                "distinct_login_days": get_optional_value(login_row, "distinct_login_days", 0),
                "active_months": get_optional_value(login_row, "active_months", 0),
                "last_login_date": get_optional_value(login_row, "last_login_date"),
                "adu_denied_attempts": adu_denied_attempts,
                "adu_denied_days": adu_denied_days,
                "last_adu_denial": get_optional_value(adu_row, "last_adu_denial"),
                "production_technician_match_status": get_optional_value(technician_row, "match_status"),
                "recommendation": recommendation,
                "rationale": rationale,
                "evidence_summary": build_recommendation_evidence_summary(
                    usage_category=usage_category,
                    distinct_login_days=get_optional_value(login_row, "distinct_login_days", 0),
                    active_months=get_optional_value(login_row, "active_months", 0),
                    adu_denied_attempts=adu_denied_attempts,
                    adu_denied_days=adu_denied_days,
                    production_match_status=get_optional_value(technician_row, "match_status"),
                ),
            }
        )

    if not rows:
        return pd.DataFrame(
            columns=[
                "user",
                "user_group",
                "usage_category",
                "distinct_login_days",
                "active_months",
                "last_login_date",
                "adu_denied_attempts",
                "adu_denied_days",
                "last_adu_denial",
                "production_technician_match_status",
                "recommendation",
                "rationale",
                "evidence_summary",
            ]
        )

    return pd.DataFrame(rows).sort_values(
        ["recommendation", "adu_denied_days", "distinct_login_days", "user"],
        ascending=[True, False, False, True],
    ).reset_index(drop=True)


def build_recommendation_evidence_summary(
    *,
    usage_category: object,
    distinct_login_days: object,
    active_months: object,
    adu_denied_attempts: int,
    adu_denied_days: int,
    production_match_status: object,
) -> str:
    """Build a concise explanation of the evidence behind a recommendation."""

    evidence = []
    if usage_category:
        evidence.append(
            f"{usage_category} PLM usage: {format_count(int(distinct_login_days or 0), 'active day')} "
            f"across {format_count(int(active_months or 0), 'active month')}"
        )
    if adu_denied_attempts:
        evidence.append(
            f"ADU access blocked {format_count(adu_denied_attempts, 'time')} "
            f"across {format_count(adu_denied_days, 'day')}"
        )
    if production_match_status == "Unmatched":
        evidence.append("No matching PLM authentication record found for Production Technician")
    return "; ".join(evidence)


def format_count(count: int, singular_label: str) -> str:
    """Format a count with a simple singular/plural label."""

    suffix = "" if count == 1 else "s"
    return f"{count} {singular_label}{suffix}"


def build_dedicated_licence_candidates(licence_recommendations: pd.DataFrame) -> pd.DataFrame:
    """Return users recommended for dedicated licence allocation with justification."""

    if licence_recommendations.empty:
        return pd.DataFrame(
            columns=[
                "user",
                "usage_category",
                "distinct_login_days",
                "active_months",
                "adu_denied_attempts",
                "adu_denied_days",
                "last_adu_denial",
                "justification",
            ]
        )

    candidates = licence_recommendations.loc[
        licence_recommendations["recommendation"] == "Consider dedicated licence allocation"
    ].copy()
    if candidates.empty:
        return pd.DataFrame(
            columns=[
                "user",
                "usage_category",
                "distinct_login_days",
                "active_months",
                "adu_denied_attempts",
                "adu_denied_days",
                "last_adu_denial",
                "justification",
            ]
        )

    candidates["justification"] = candidates["evidence_summary"].apply(
        lambda value: (
            f"{value}. Provision is justified because the user has observed PLM usage and repeated failed "
            "attempts to obtain an ADU/shared licence."
        )
    )
    return candidates[
        [
            "user",
            "usage_category",
            "distinct_login_days",
            "active_months",
            "adu_denied_attempts",
            "adu_denied_days",
            "last_adu_denial",
            "justification",
        ]
    ].sort_values(["adu_denied_days", "adu_denied_attempts", "distinct_login_days"], ascending=[False, False, False])


def build_unused_licence_evidence(licence_recommendations: pd.DataFrame) -> pd.DataFrame:
    """Build the evidence table showing recoverable licences after proposed allocations."""

    if licence_recommendations.empty:
        recoverable = 0
        dedicated_demand = 0
    else:
        recoverable = int(
            (licence_recommendations["recommendation"] == "Candidate for licence removal/reallocation").sum()
        )
        dedicated_demand = int(
            (licence_recommendations["recommendation"] == "Consider dedicated licence allocation").sum()
        )
    remaining = recoverable - dedicated_demand
    return pd.DataFrame(
        [
            {
                "measure": "Potential recoverable Production Technician licences",
                "count": recoverable,
                "confirmation_method": (
                    "Count of Production Technicians either unmatched to the PLM login audit "
                    "or matched but classified as Rare usage."
                ),
            },
            {
                "measure": "Dedicated licences proposed for ADU users",
                "count": dedicated_demand,
                "confirmation_method": (
                    "Count of ADU-denied users with observed Regular or Occasional PLM usage and ADU denied days "
                    "above the dedicated-licence threshold."
                ),
            },
            {
                "measure": "Recoverable licences remaining after proposed allocations",
                "count": remaining,
                "confirmation_method": (
                    "Potential recoverable Production Technician licences minus proposed dedicated ADU allocations."
                ),
            },
        ]
    )


def choose_recommendation_user_name(
    login_row: pd.Series | None,
    adu_row: pd.Series | None,
    technician_row: pd.Series | None,
) -> str:
    """Choose the clearest display name for a recommendation row."""

    for row, column in (
        (login_row, "user_display_name"),
        (adu_row, "user_display_name"),
        (technician_row, "production_technician_name"),
    ):
        value = get_optional_value(row, column)
        if value:
            return str(value)
    return ""


def build_user_group_label(is_production_technician: bool, has_adu_denials: bool) -> str:
    """Return the combined evidence group for a recommendation row."""

    groups = []
    if is_production_technician:
        groups.append("Production Technician")
    if has_adu_denials:
        groups.append("ADU Denied")
    return " + ".join(groups)


def classify_licence_recommendation(
    *,
    usage_category: object,
    is_production_technician: bool,
    production_match_status: object,
    has_adu_denials: bool,
    adu_denied_attempts: int,
    adu_denied_days: int,
    config: AppConfig,
) -> tuple[str, str]:
    """Classify a user into an actionable licence recommendation."""

    repeated_adu_denials = (
        adu_denied_days >= config.adu_repeated_denial_days_threshold
        or adu_denied_attempts >= config.adu_repeated_denial_attempts_threshold
    )
    high_confidence_adu_demand = adu_denied_days >= config.dedicated_licence_denied_days_threshold
    if is_production_technician and production_match_status != "Matched":
        return (
            "Candidate for licence removal/reallocation",
            "No matching PLM authentication record found for this Production Technician in the audit period.",
        )
    if is_production_technician and usage_category == "Rare":
        return (
            "Candidate for licence removal/reallocation",
            "Production Technician matched to PLM but classified as Rare usage.",
        )
    if is_production_technician and usage_category == "Occasional":
        return ("Review business need", "Production Technician has occasional observed PLM usage.")
    if is_production_technician and usage_category == "Regular":
        return ("Retain dedicated licence", "Production Technician has regular observed PLM usage.")
    if has_adu_denials and high_confidence_adu_demand and usage_category in {"Regular", "Occasional"}:
        return (
            "Consider dedicated licence allocation",
            "User has observed PLM usage and ADU denied days above the dedicated-licence threshold.",
        )
    if has_adu_denials and repeated_adu_denials:
        return (
            "Review licence need",
            "User has repeated ADU denial evidence but is below the high-confidence dedicated-licence threshold.",
        )
    if has_adu_denials:
        return ("Monitor", "User has ADU denial evidence but not repeated enough for automatic escalation.")
    return ("No action", "No licence review trigger found.")


def get_optional_value(row: pd.Series | None, column: str, default=None):
    """Read a value from an optional pandas row."""

    if row is None or column not in row:
        return default
    value = row[column]
    if pd.isna(value):
        return default
    return value


def build_recommendation_summary(licence_recommendations: pd.DataFrame) -> pd.DataFrame:
    """Summarise recommendation counts."""

    if licence_recommendations.empty:
        return pd.DataFrame(columns=["recommendation", "user_count"])
    return licence_recommendations.groupby("recommendation").size().rename("user_count").reset_index().sort_values(
        "user_count",
        ascending=False,
    ).reset_index(drop=True)


def build_licence_balance(licence_recommendations: pd.DataFrame) -> pd.DataFrame:
    """Build a management-level view of potential recovery versus demand."""

    if licence_recommendations.empty:
        recoverable = 0
        demand = 0
    else:
        recoverable = int(
            (licence_recommendations["recommendation"] == "Candidate for licence removal/reallocation").sum()
        )
        demand = int(
            (licence_recommendations["recommendation"] == "Consider dedicated licence allocation").sum()
        )
    return pd.DataFrame(
        [
            {"licence_position": "Potential licences recoverable", "user_count": recoverable},
            {"licence_position": "Potential dedicated licence demand", "user_count": demand},
        ]
    )


def build_executive_summary(
    cleaned_df: pd.DataFrame,
    user_summary: pd.DataFrame,
    adu_user_summary: pd.DataFrame,
    production_technician_matches: pd.DataFrame,
    licence_recommendations: pd.DataFrame,
) -> pd.DataFrame:
    """Build headline story metrics for the Executive Summary sheet."""

    production_no_auth = 0
    if not production_technician_matches.empty:
        production_no_auth = int((production_technician_matches["match_status"] != "Matched").sum())

    recoverable = 0
    dedicated_demand = 0
    if not licence_recommendations.empty:
        recoverable = int(
            (licence_recommendations["recommendation"] == "Candidate for licence removal/reallocation").sum()
        )
        dedicated_demand = int(
            (licence_recommendations["recommendation"] == "Consider dedicated licence allocation").sum()
        )
    remaining_recoverable = recoverable - dedicated_demand

    metrics = [
        ("PLM login records analysed", len(cleaned_df)),
        ("PLM users analysed", user_summary["user"].nunique()),
        ("Regular PLM users", int((user_summary["usage_category"] == "Regular").sum())),
        ("Occasional PLM users", int((user_summary["usage_category"] == "Occasional").sum())),
        ("Rare PLM users", int((user_summary["usage_category"] == "Rare").sum())),
        ("Users with ADU denial evidence", len(adu_user_summary)),
        ("Production Technicians reviewed", len(production_technician_matches)),
        ("Production Technicians with no observed PLM authentication", production_no_auth),
        ("Potential licences recoverable", recoverable),
        ("Potential dedicated licence demand", dedicated_demand),
        ("Recoverable licences remaining after proposed allocations", remaining_recoverable),
    ]
    return pd.DataFrame(metrics, columns=["metric", "value"])


def build_most_active_users(user_summary: pd.DataFrame, limit: int = 50) -> pd.DataFrame:
    """Build a table of the most active users for reporting charts."""

    most_active = user_summary.sort_values(
        [
            "average_active_days_per_month",
            "distinct_login_days",
            "total_logins",
            "user_display_name",
            "user",
        ],
        ascending=[False, False, False, True, True],
    ).head(limit)

    return most_active[
        [
            "user_display_name",
            "usage_category",
            "average_active_days_per_month",
            "distinct_login_days",
            "total_logins",
            "last_login_date",
        ]
    ].reset_index(drop=True)


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
