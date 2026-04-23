"""Production Technician list loading and matching helpers."""

from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

from config import AppConfig

LOGGER = logging.getLogger(__name__)


def load_production_technicians(config: AppConfig) -> pd.DataFrame:
    """Load the optional Production Technician full-name list."""

    if config.production_technicians_file is None:
        return build_empty_production_technician_list()

    path = config.production_technicians_file
    if not path.exists():
        raise FileNotFoundError(f"Production Technician file not found: {path}")

    LOGGER.info("Loading Production Technician list from %s", path)
    dataframe = read_production_technician_file(path, config)
    dataframe.columns = [str(column).strip() for column in dataframe.columns]

    if config.production_technicians_name_column not in dataframe.columns:
        raise ValueError(
            "Missing Production Technician name column: "
            f"{config.production_technicians_name_column}"
        )

    names = dataframe[config.production_technicians_name_column].astype("string").str.strip()
    output = pd.DataFrame({"production_technician_name": names})
    output = output.loc[output["production_technician_name"].notna()]
    output = output.loc[output["production_technician_name"] != ""].copy()
    output["production_technician_match_key"] = output["production_technician_name"].apply(
        normalise_person_name
    )
    output = output.drop_duplicates("production_technician_match_key").sort_values(
        "production_technician_name"
    )
    return output.reset_index(drop=True)


def read_production_technician_file(path: Path, config: AppConfig) -> pd.DataFrame:
    """Read a Production Technician source file from Excel or CSV."""

    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, sheet_name=config.production_technicians_sheet)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    raise ValueError(f"Unsupported Production Technician file type: {path.suffix}")


def build_empty_production_technician_list() -> pd.DataFrame:
    """Return an empty Production Technician list with the expected columns."""

    return pd.DataFrame(
        columns=[
            "production_technician_name",
            "production_technician_match_key",
        ]
    )


def normalise_person_name(value: object) -> str:
    """Normalise a person's full name for conservative exact matching."""

    if pd.isna(value):
        return ""
    text = str(value).strip().casefold()
    text = re.sub(r"\s+", " ", text)
    return text


def build_production_technician_match_report(
    production_technicians: pd.DataFrame,
    user_summary: pd.DataFrame,
) -> pd.DataFrame:
    """Build a match-quality report for the Production Technician list."""

    if production_technicians.empty:
        return pd.DataFrame(
            columns=[
                "production_technician_name",
                "match_status",
                "matched_user",
                "usage_category",
                "review_status",
                "distinct_login_days",
                "last_login_date",
            ]
        )

    users = user_summary.copy()
    users["production_technician_match_key"] = users["user_display_name"].apply(normalise_person_name)
    matched = production_technicians.merge(
        users[
            [
                "production_technician_match_key",
                "user",
                "user_display_name",
                "usage_category",
                "distinct_login_days",
                "last_login_date",
            ]
        ],
        on="production_technician_match_key",
        how="left",
    )
    matched["match_status"] = matched["user"].apply(lambda value: "Matched" if pd.notna(value) else "Unmatched")
    matched = matched.rename(columns={"user": "matched_user"})
    matched["review_status"] = matched.apply(classify_review_status, axis=1)
    return matched[
        [
            "production_technician_name",
            "match_status",
            "matched_user",
            "user_display_name",
            "usage_category",
            "review_status",
            "distinct_login_days",
            "last_login_date",
        ]
    ].sort_values(["match_status", "production_technician_name"]).reset_index(drop=True)


def classify_review_status(row: pd.Series) -> str:
    """Classify Production Technician licence review status."""

    if row["match_status"] != "Matched":
        return "No observed PLM authentication"
    if row["usage_category"] == "Regular":
        return "Retain dedicated licence"
    if row["usage_category"] == "Occasional":
        return "Review business need"
    return "Candidate for licence removal/reallocation"
