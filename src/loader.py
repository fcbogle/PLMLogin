"""Excel loading utilities."""

from __future__ import annotations

import logging

import pandas as pd

from config import AppConfig

LOGGER = logging.getLogger(__name__)


def load_login_records(config: AppConfig) -> pd.DataFrame:
    """Load login records from the configured Excel workbook and sheet."""

    if not config.input_file.exists():
        raise FileNotFoundError(f"Input file not found: {config.input_file}")

    LOGGER.info(
        "Loading input workbook %s [sheet=%s]",
        config.input_file,
        config.input_sheet,
    )
    dataframe = pd.read_excel(config.input_file, sheet_name=config.input_sheet)
    dataframe.columns = [str(column).strip() for column in dataframe.columns]
    return dataframe
