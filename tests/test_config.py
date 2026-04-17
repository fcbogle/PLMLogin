"""Tests for configuration helpers."""

from __future__ import annotations

from pathlib import Path

import pytest

from config import build_config_with_overrides, load_exclusion_file


def test_load_exclusion_file_reads_exact_and_contains_values(tmp_path: Path) -> None:
    path = tmp_path / "exclusions.json"
    path.write_text(
        (
            "{\n"
            '  "excluded_user_exact_values": ["Exact User"],\n'
            '  "excluded_user_contains_values": ["service"]\n'
            "}\n"
        ),
        encoding="utf-8",
    )

    exclusions = load_exclusion_file(path)

    assert exclusions == {
        "exact_values": ("Exact User",),
        "contains_values": ("service",),
    }


def test_load_exclusion_file_rejects_invalid_payload(tmp_path: Path) -> None:
    path = tmp_path / "exclusions.json"
    path.write_text('{"excluded_user_exact_values": "not-a-list"}', encoding="utf-8")

    with pytest.raises(ValueError, match="excluded_user_exact_values"):
        load_exclusion_file(path)


def test_build_config_with_overrides_can_disable_defaults_and_use_file_values(tmp_path: Path) -> None:
    path = tmp_path / "exclusions.json"
    path.write_text(
        (
            "{\n"
            '  "excluded_user_exact_values": ["Custom Exact"],\n'
            '  "excluded_user_contains_values": ["custom-fragment"]\n'
            "}\n"
        ),
        encoding="utf-8",
    )

    config = build_config_with_overrides(
        exclusions_file=path,
        disable_default_exclusions=True,
    )

    assert config.excluded_user_exact_values == ("Custom Exact",)
    assert config.excluded_user_contains_values == ("custom-fragment",)
