"""Tests for usage categorisation."""

from __future__ import annotations

from config import CategoryThresholds
from src.categoriser import categorise_user


def test_categorise_user_uses_threshold_boundaries() -> None:
    thresholds = CategoryThresholds(regular_min_average_active_days=8.0, occasional_min_average_active_days=2.0)

    assert categorise_user(8.0, thresholds) == "Regular"
    assert categorise_user(7.99, thresholds) == "Occasional"
    assert categorise_user(2.0, thresholds) == "Occasional"
    assert categorise_user(1.99, thresholds) == "Rare"
