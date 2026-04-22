"""Tests for Production Technician matching."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from config import AppConfig
from src.analyser import add_production_technician_flags
from src.production_technicians import (
    build_production_technician_match_report,
    load_production_technicians,
    normalise_person_name,
)


def test_normalise_person_name_is_case_and_spacing_insensitive() -> None:
    assert normalise_person_name("  Andy   Self ") == "andy self"
    assert normalise_person_name("ANDY SELF") == "andy self"


def test_load_production_technicians_reads_csv_and_deduplicates(tmp_path: Path) -> None:
    path = tmp_path / "production_technicians.csv"
    path.write_text("Full Name\nAndy Self\nandy  self\nBeth Jones\n\n", encoding="utf-8")

    technicians = load_production_technicians(AppConfig(production_technicians_file=path))

    assert technicians["production_technician_name"].tolist() == ["Andy Self", "Beth Jones"]
    assert technicians["production_technician_match_key"].tolist() == ["andy self", "beth jones"]


def test_add_production_technician_flags_matches_user_display_name() -> None:
    user_summary = pd.DataFrame(
        {
            "user": ["Andy Self (Andy Self: Blatchford)", "Cara Smith (Cara Smith: Blatchford)"],
            "user_display_name": ["Andy Self", "Cara Smith"],
            "usage_category": ["Rare", "Regular"],
        }
    )
    technicians = pd.DataFrame(
        {
            "production_technician_name": ["andy  self"],
            "production_technician_match_key": ["andy self"],
        }
    )

    flagged = add_production_technician_flags(user_summary, technicians)

    assert flagged["is_production_technician"].tolist() == [True, False]
    assert flagged["production_technician_name"].tolist()[0] == "andy  self"


def test_build_production_technician_match_report_flags_unmatched_names() -> None:
    technicians = pd.DataFrame(
        {
            "production_technician_name": ["Andy Self", "Beth Jones"],
            "production_technician_match_key": ["andy self", "beth jones"],
        }
    )
    user_summary = pd.DataFrame(
        {
            "user": ["Andy Self (Andy Self: Blatchford)"],
            "user_display_name": ["Andy Self"],
            "usage_category": ["Rare"],
            "distinct_login_days": [3],
            "last_login_date": [pd.Timestamp("2026-04-20")],
        }
    )

    report = build_production_technician_match_report(technicians, user_summary)

    assert report["match_status"].tolist() == ["Matched", "Unmatched"]
    assert report["production_technician_name"].tolist() == ["Andy Self", "Beth Jones"]
