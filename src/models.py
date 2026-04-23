"""Shared dataclasses used across the application."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class CleaningReport:
    """Describes how the raw dataset was cleaned."""

    input_rows: int
    output_rows: int
    dropped_missing_user_rows: int
    dropped_invalid_timestamp_rows: int
    dropped_excluded_user_rows: int


@dataclass(frozen=True)
class AnalysisOutputs:
    """Container for all tabular outputs required by the workbook."""

    raw_data: pd.DataFrame
    user_summary: pd.DataFrame
    regular_users: pd.DataFrame
    occasional_users: pd.DataFrame
    rare_users: pd.DataFrame
    adu_raw_data: pd.DataFrame
    adu_user_summary: pd.DataFrame
    adu_monthly_denials: pd.DataFrame
    monthly_activity: pd.DataFrame
    licence_recommendations: pd.DataFrame
    dedicated_licence_candidates: pd.DataFrame
    unused_licence_evidence: pd.DataFrame
    executive_summary: pd.DataFrame
    recommendation_summary: pd.DataFrame
    production_technician_review_summary: pd.DataFrame
    licence_balance: pd.DataFrame
    overview_metrics: pd.DataFrame
    category_summary: pd.DataFrame
    monthly_active_users: pd.DataFrame
    most_active_users: pd.DataFrame
    at_risk_rare_users: pd.DataFrame
    production_technician_matches: pd.DataFrame
    category_rules: pd.DataFrame
    cleaning_report: CleaningReport
