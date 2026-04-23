"""Named licence allocation loading and matching helpers."""

from __future__ import annotations

import logging

import pandas as pd

from config import AppConfig
from src.production_technicians import normalise_person_name

LOGGER = logging.getLogger(__name__)


def load_named_licences(config: AppConfig) -> pd.DataFrame:
    """Load the optional named licence assignment workbook."""

    if config.named_licence_input_file is None:
        return build_empty_named_licences()
    if not config.named_licence_input_file.exists():
        raise FileNotFoundError(f"Named licence input file not found: {config.named_licence_input_file}")

    LOGGER.info(
        "Loading named licence workbook %s [sheet=%s]",
        config.named_licence_input_file,
        config.named_licence_input_sheet,
    )
    dataframe = pd.read_excel(config.named_licence_input_file, sheet_name=config.named_licence_input_sheet)
    dataframe.columns = [str(column).strip() for column in dataframe.columns]

    required_columns = {config.named_licence_user_column, config.named_licence_allocated_column}
    missing_columns = required_columns - set(dataframe.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Missing required named licence columns: {missing}")

    output = dataframe[
        [
            config.named_licence_user_column,
            config.named_licence_allocated_column,
        ]
    ].rename(
        columns={
            config.named_licence_user_column: "named_user",
            config.named_licence_allocated_column: "allocated_licence",
        }
    )
    output["named_user"] = output["named_user"].astype("string").str.strip()
    output["allocated_licence"] = output["allocated_licence"].astype("string").str.strip()
    output = output.loc[output["named_user"].notna() & (output["named_user"] != "")].copy()
    output["named_user_match_key"] = output["named_user"].apply(normalise_person_name)
    return output.drop_duplicates("named_user_match_key").sort_values("named_user").reset_index(drop=True)


def build_empty_named_licences() -> pd.DataFrame:
    """Return an empty named licence dataframe with expected columns."""

    return pd.DataFrame(columns=["named_user", "allocated_licence", "named_user_match_key"])
