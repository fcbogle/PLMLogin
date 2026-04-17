"""User categorisation rules."""

from __future__ import annotations

from config import CategoryThresholds


def categorise_user(average_active_days_per_month: float, thresholds: CategoryThresholds) -> str:
    """Return the usage category for a user."""

    if average_active_days_per_month >= thresholds.regular_min_average_active_days:
        return "Regular"
    if average_active_days_per_month >= thresholds.occasional_min_average_active_days:
        return "Occasional"
    return "Rare"
