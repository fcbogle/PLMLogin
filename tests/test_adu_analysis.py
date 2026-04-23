"""Tests for ADU denial analysis and licence recommendations."""

from __future__ import annotations

import pandas as pd

from config import AppConfig
from src.adu import build_adu_monthly_denials, build_adu_user_summary, clean_adu_denials
from src.analyser import build_licence_recommendations


def test_clean_adu_denials_and_build_user_summary() -> None:
    raw_df = pd.DataFrame(
        {
            "Event Label": [
                "Login Denied - Insufficient ADU License",
                "Login Denied - Insufficient ADU License",
                "Other Event",
            ],
            "Event Time": [
                "2026-04-15 11:00:00 BST",
                "2026-04-16 12:00:00 BST",
                "2026-04-16 12:30:00 BST",
            ],
            "User Name": [
                "Ada Lovelace (Ada Lovelace: Blatchford)",
                "Ada Lovelace (Ada Lovelace: Blatchford)",
                "Grace Hopper (Grace Hopper: Blatchford)",
            ],
        }
    )

    cleaned = clean_adu_denials(raw_df, AppConfig(excluded_user_exact_values=(), exclusions_file=None))
    summary = build_adu_user_summary(cleaned, AppConfig())
    monthly = build_adu_monthly_denials(cleaned)

    assert len(cleaned) == 2
    assert summary.loc[0, "user_display_name"] == "Ada Lovelace"
    assert summary.loc[0, "adu_denied_attempts"] == 2
    assert summary.loc[0, "adu_denied_days"] == 2
    assert bool(summary.loc[0, "is_repeated_adu_denial_user"]) is True
    assert monthly.loc[0, "adu_denied_attempts"] == 2


def test_licence_recommendations_combine_adu_and_production_technicians() -> None:
    user_summary = pd.DataFrame(
        {
            "user": ["Ada Lovelace (Ada Lovelace: Blatchford)"],
            "user_display_name": ["Ada Lovelace"],
            "usage_category": ["Regular"],
            "distinct_login_days": [120],
            "active_months": [12],
            "last_login_date": [pd.Timestamp("2026-04-16")],
        }
    )
    adu_summary = pd.DataFrame(
        {
            "user": ["Ada Lovelace (Ada Lovelace: Blatchford)"],
            "user_display_name": ["Ada Lovelace"],
            "adu_denied_attempts": [15],
            "adu_denied_days": [10],
            "last_adu_denial": [pd.Timestamp("2026-04-16")],
        }
    )
    production_matches = pd.DataFrame(
        {
            "production_technician_name": ["Andy Self"],
            "match_status": ["Unmatched"],
            "matched_user": [pd.NA],
            "user_display_name": [pd.NA],
            "usage_category": [pd.NA],
            "review_status": ["No observed PLM authentication"],
            "distinct_login_days": [pd.NA],
            "last_login_date": [pd.NaT],
        }
    )

    recommendations = build_licence_recommendations(
        user_summary,
        adu_summary,
        production_matches,
        AppConfig(),
    )

    by_user = recommendations.set_index("user")
    assert by_user.loc["Ada Lovelace", "recommendation"] == "Consider dedicated licence allocation"
    assert by_user.loc["Andy Self", "recommendation"] == "Candidate for licence removal/reallocation"
