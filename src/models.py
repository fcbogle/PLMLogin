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
    monthly_activity: pd.DataFrame
    overview_metrics: pd.DataFrame
    category_summary: pd.DataFrame
    category_rules: pd.DataFrame
    cleaning_report: CleaningReport
