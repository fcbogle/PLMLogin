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
    output_file: Path = Path("output/plm_usage_and_licence_optimisation_analysis.xlsx")
    user_column: str = "User Name"
    timestamp_column: str = "Event Time"
    adu_input_file: Path | None = Path("/Users/frankbogle/Documents/adu/Insufficient ADU License-Login Audit Repot.xlsx")
    adu_input_sheet: str = "AuditReportExport (1)"
    adu_user_column: str = "User Name"
    adu_timestamp_column: str = "Event Time"
    adu_event_label_column: str = "Event Label"
    adu_denied_event_label: str = "Login Denied - Insufficient ADU License"
    named_licence_input_file: Path | None = Path("/Users/frankbogle/Documents/act_license/Licensed_User-23-04-26.xlsx")
    named_licence_input_sheet: str = "Sheet1"
    named_licence_user_column: str = "User"
    named_licence_allocated_column: str = "Allocated License"
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
    production_technicians_file: Path | None = None
    production_technicians_sheet: str | int = 0
    production_technicians_name_column: str = "Full Name"
    thresholds: CategoryThresholds = field(default_factory=CategoryThresholds)
    recent_activity_days: int = 90
    adu_repeated_denial_days_threshold: int = 2
    adu_repeated_denial_attempts_threshold: int = 3
    dedicated_licence_denied_days_threshold: int = 10


DEFAULT_CONFIG = AppConfig()


def build_config_with_overrides(
    *,
    input_file: Path | None = None,
    input_sheet: str | None = None,
    output_file: Path | None = None,
    user_column: str | None = None,
    timestamp_column: str | None = None,
    adu_input_file: Path | None = None,
    adu_input_sheet: str | None = None,
    adu_user_column: str | None = None,
    adu_timestamp_column: str | None = None,
    adu_event_label_column: str | None = None,
    named_licence_input_file: Path | None = None,
    named_licence_input_sheet: str | None = None,
    named_licence_user_column: str | None = None,
    named_licence_allocated_column: str | None = None,
    exclusions_file: Path | None = None,
    production_technicians_file: Path | None = None,
    production_technicians_sheet: str | int | None = None,
    production_technicians_name_column: str | None = None,
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
        adu_input_file=adu_input_file if adu_input_file is not None else config.adu_input_file,
        adu_input_sheet=adu_input_sheet or config.adu_input_sheet,
        adu_user_column=adu_user_column or config.adu_user_column,
        adu_timestamp_column=adu_timestamp_column or config.adu_timestamp_column,
        adu_event_label_column=adu_event_label_column or config.adu_event_label_column,
        named_licence_input_file=(
            named_licence_input_file if named_licence_input_file is not None else config.named_licence_input_file
        ),
        named_licence_input_sheet=named_licence_input_sheet or config.named_licence_input_sheet,
        named_licence_user_column=named_licence_user_column or config.named_licence_user_column,
        named_licence_allocated_column=(
            named_licence_allocated_column or config.named_licence_allocated_column
        ),
        production_technicians_file=production_technicians_file or config.production_technicians_file,
        production_technicians_sheet=(
            config.production_technicians_sheet
            if production_technicians_sheet is None
            else production_technicians_sheet
        ),
        production_technicians_name_column=(
            production_technicians_name_column or config.production_technicians_name_column
        ),
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
