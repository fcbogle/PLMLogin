"""CLI entrypoint for PLM login analysis."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path

from config import build_config_with_overrides
from src.analyser import build_analysis_outputs
from src.cleaner import clean_login_data
from src.excel_writer import write_analysis_workbook
from src.loader import load_login_records
from src.production_technicians import load_production_technicians


def configure_logging() -> None:
    """Configure basic console logging."""

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    )


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""

    parser = argparse.ArgumentParser(description="Analyse PLM login activity from an Excel workbook.")
    parser.add_argument("--input-file", type=Path, help="Path to the source Excel workbook.")
    parser.add_argument("--input-sheet", help="Name of the source worksheet to analyse.")
    parser.add_argument("--output-file", type=Path, help="Path for the generated Excel workbook.")
    parser.add_argument("--user-column", help="Column name containing the user identifier.")
    parser.add_argument("--timestamp-column", help="Column name containing the login timestamp.")
    parser.add_argument(
        "--exclusions-file",
        type=Path,
        help="Path to a JSON file containing user exclusion values.",
    )
    parser.add_argument(
        "--normalise-user-case",
        action="store_true",
        help="Convert user identifiers to lower case during cleaning.",
    )
    parser.add_argument(
        "--disable-default-exclusions",
        action="store_true",
        help="Include the built-in test/admin accounts in the analysis.",
    )
    parser.add_argument(
        "--production-technicians-file",
        type=Path,
        help="Optional Excel or CSV file containing Production Technician full names.",
    )
    parser.add_argument(
        "--production-technicians-sheet",
        help="Worksheet name for the Production Technician file. Defaults to the first sheet.",
    )
    parser.add_argument(
        "--production-technicians-name-column",
        help='Column containing Production Technician full names. Defaults to "Full Name".',
    )
    return parser.parse_args()


def main() -> None:
    """Run the end-to-end PLM login analysis workflow."""

    configure_logging()
    args = parse_args()
    config = build_config_with_overrides(
        input_file=args.input_file,
        input_sheet=args.input_sheet,
        output_file=args.output_file,
        user_column=args.user_column,
        timestamp_column=args.timestamp_column,
        exclusions_file=args.exclusions_file,
        production_technicians_file=args.production_technicians_file,
        production_technicians_sheet=args.production_technicians_sheet,
        production_technicians_name_column=args.production_technicians_name_column,
        normalise_user_case=True if args.normalise_user_case else None,
        disable_default_exclusions=args.disable_default_exclusions,
    )

    raw_df = load_login_records(config)
    cleaned_df, cleaning_report = clean_login_data(raw_df, config)
    production_technicians = load_production_technicians(config)
    outputs = build_analysis_outputs(cleaned_df, config, cleaning_report, production_technicians)
    write_analysis_workbook(outputs, config)

    logging.info("Analysis workbook created at %s", config.output_file.resolve())


if __name__ == "__main__":
    main()
