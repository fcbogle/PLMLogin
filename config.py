"""Application configuration for PLM login analysis."""

from __future__ import annotations

import json
from dataclasses import dataclass, field, replace
from pathlib import Path


@dataclass(frozen=True)
class CategoryThresholds:
    """Thresholds used to categorise users by usage regularity."""

    regular_min_average_active_days: float = 8.0
    occasional_min_average_active_days: float = 2.0


@dataclass(frozen=True)
class AppConfig:
    """Runtime configuration for the analysis pipeline."""

    input_file: Path = Path("/Users/frankbogle/Documents/login/BlatchfordUserLoginAuditReportExport.xlsx")
    input_sheet: str = "AuditReportExport"
    output_file: Path = Path("output/plm_login_analysis.xlsx")
    user_column: str = "User Name"
    timestamp_column: str = "Event Time"
    optional_columns: list[str] = field(
        default_factory=lambda: [
            "Event Label",
            "IP Address",
            "User Organization",
            "Unique User Names",
        ]
    )
    normalise_user_case: bool = False
    timestamp_suffixes_to_strip: tuple[str, ...] = (" BST", " GMT")
    excluded_user_exact_values: tuple[str, ...] = (
        "WPS Test (WPS Test: Blatchford)",
        "PLM Test (PLM Test: Blatchford)",
        "orgadmin (orgadmin: Blatchford)",
        "bfpubadmin (bfpubadmin: Blatchford)",
        "bfwcadmin (bfwcadmin: Blatchford)",
        "PLM C (PLM C: Blatchford)",
        "Site Administrator (SiteAdmin: Site)",
    )
    excluded_user_contains_values: tuple[str, ...] = ()
    exclusions_file: Path = Path("config/exclusions.json")
    thresholds: CategoryThresholds = field(default_factory=CategoryThresholds)
    recent_activity_days: int = 90


DEFAULT_CONFIG = AppConfig()


def build_config_with_overrides(
    *,
    input_file: Path | None = None,
    input_sheet: str | None = None,
    output_file: Path | None = None,
    user_column: str | None = None,
    timestamp_column: str | None = None,
    exclusions_file: Path | None = None,
    normalise_user_case: bool | None = None,
    disable_default_exclusions: bool = False,
) -> AppConfig:
    """Return the default config with optional runtime overrides applied."""

    config = DEFAULT_CONFIG
    exclusion_file_path = exclusions_file or config.exclusions_file
    file_exclusions = load_exclusion_file(exclusion_file_path)
    if disable_default_exclusions:
        config = replace(
            config,
            excluded_user_exact_values=(),
            excluded_user_contains_values=(),
        )

    return replace(
        config,
        input_file=input_file or config.input_file,
        input_sheet=input_sheet or config.input_sheet,
        output_file=output_file or config.output_file,
        user_column=user_column or config.user_column,
        timestamp_column=timestamp_column or config.timestamp_column,
        excluded_user_exact_values=config.excluded_user_exact_values + file_exclusions["exact_values"],
        excluded_user_contains_values=config.excluded_user_contains_values + file_exclusions["contains_values"],
        exclusions_file=exclusion_file_path,
        normalise_user_case=(
            config.normalise_user_case if normalise_user_case is None else normalise_user_case
        ),
    )


def load_exclusion_file(path: Path) -> dict[str, tuple[str, ...]]:
    """Load user exclusion values from a JSON file if it exists."""

    if not path.exists():
        return {"exact_values": (), "contains_values": ()}

    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    if not isinstance(payload, dict):
        raise ValueError(f"Exclusions file must contain a JSON object: {path}")

    exact_values = validate_string_list(payload.get("excluded_user_exact_values", []), path, "excluded_user_exact_values")
    contains_values = validate_string_list(
        payload.get("excluded_user_contains_values", []),
        path,
        "excluded_user_contains_values",
    )
    return {
        "exact_values": tuple(exact_values),
        "contains_values": tuple(contains_values),
    }


def validate_string_list(value: object, path: Path, field_name: str) -> list[str]:
    """Validate that a JSON field is a list of strings."""

    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError(f"{field_name} in {path} must be a JSON array of strings")
    return value
