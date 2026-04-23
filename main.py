"""CLI entrypoint for PLM login analysis."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path

from config import build_config_with_overrides
from src.adu import build_adu_monthly_denials, build_adu_user_summary, clean_adu_denials, load_adu_denials
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

    parser = argparse.ArgumentParser(
        description="Analyse PLM usage, ADU denials, and licence review evidence."
    )
    parser.add_argument("--input-file", type=Path, help="Path to the source Excel workbook.")
    parser.add_argument("--input-sheet", help="Name of the source worksheet to analyse.")
    parser.add_argument("--output-file", type=Path, help="Path for the generated Excel workbook.")
    parser.add_argument("--user-column", help="Column name containing the user identifier.")
    parser.add_argument("--timestamp-column", help="Column name containing the login timestamp.")
    parser.add_argument("--adu-input-file", type=Path, help="Path to the ADU denial audit workbook.")
    parser.add_argument("--adu-input-sheet", help="Name of the ADU denial audit worksheet.")
    parser.add_argument("--adu-user-column", help="ADU audit column containing the user identifier.")
    parser.add_argument("--adu-timestamp-column", help="ADU audit column containing the event timestamp.")
    parser.add_argument("--adu-event-label-column", help="ADU audit column containing the event label.")
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
        adu_input_file=args.adu_input_file,
        adu_input_sheet=args.adu_input_sheet,
        adu_user_column=args.adu_user_column,
        adu_timestamp_column=args.adu_timestamp_column,
        adu_event_label_column=args.adu_event_label_column,
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
    raw_adu_df = load_adu_denials(config)
    cleaned_adu_df = clean_adu_denials(raw_adu_df, config)
    adu_user_summary = build_adu_user_summary(cleaned_adu_df, config)
    adu_monthly_denials = build_adu_monthly_denials(cleaned_adu_df)
    outputs = build_analysis_outputs(
        cleaned_df,
        config,
        cleaning_report,
        production_technicians,
        cleaned_adu_df,
        adu_user_summary,
        adu_monthly_denials,
    )
    write_analysis_workbook(outputs, config)

    logging.info("Analysis workbook created at %s", config.output_file.resolve())


if __name__ == "__main__":
    main()
