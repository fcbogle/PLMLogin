"""Workbook writer and formatting for analysis outputs."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import CellIsRule

from config import AppConfig
from src.models import AnalysisOutputs


HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9E2F3")
CATEGORY_FILLS = {
    "Regular": PatternFill(fill_type="solid", fgColor="C6EFCE"),
    "Occasional": PatternFill(fill_type="solid", fgColor="FFEB9C"),
    "Rare": PatternFill(fill_type="solid", fgColor="F4CCCC"),
}


def write_analysis_workbook(outputs: AnalysisOutputs, config: AppConfig) -> None:
    """Write all required sheets to a formatted Excel workbook."""

    output_path = config.output_file
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    workbook.remove(workbook.active)

    create_overview_sheet(workbook.create_sheet("Overview"), outputs)
    write_table_sheet(workbook.create_sheet("Raw_Data"), outputs.raw_data)
    write_table_sheet(workbook.create_sheet("User_Summary"), outputs.user_summary, category_column="usage_category")
    write_table_sheet(workbook.create_sheet("Regular_Users"), outputs.regular_users, category_column="usage_category")
    write_table_sheet(
        workbook.create_sheet("Occasional_Users"),
        outputs.occasional_users,
        category_column="usage_category",
    )
    write_table_sheet(workbook.create_sheet("Rare_Users"), outputs.rare_users, category_column="usage_category")
    write_table_sheet(workbook.create_sheet("Monthly_Activity"), outputs.monthly_activity)
    write_table_sheet(workbook.create_sheet("Category_Rules"), outputs.category_rules)

    workbook.save(output_path)


def create_overview_sheet(worksheet: Worksheet, outputs: AnalysisOutputs) -> None:
    """Create the overview sheet with metrics and category summary."""

    worksheet.append(["Metric", "Value"])
    for row in outputs.overview_metrics.itertuples(index=False):
        worksheet.append(list(row))

    worksheet["D1"] = "Category"
    worksheet["E1"] = "User Count"
    worksheet["F1"] = "Percentage of Total Users"
    for row_index, row in enumerate(outputs.category_summary.itertuples(index=False), start=2):
        worksheet.cell(row=row_index, column=4, value=row.category)
        worksheet.cell(row=row_index, column=5, value=row.user_count)
        worksheet.cell(row=row_index, column=6, value=row.percentage_of_total_users)

    format_sheet(worksheet)
    worksheet.freeze_panes = "A2"


def write_table_sheet(
    worksheet: Worksheet,
    dataframe,
    category_column: str | None = None,
) -> None:
    """Write a dataframe to a worksheet and apply basic formatting."""

    worksheet.append(list(dataframe.columns))
    for row in dataframe.itertuples(index=False, name=None):
        worksheet.append(list(row))

    format_sheet(worksheet)
    if category_column and category_column in dataframe.columns:
        apply_category_highlighting(worksheet, dataframe.columns.get_loc(category_column) + 1)


def format_sheet(worksheet: Worksheet) -> None:
    """Apply standard workbook formatting."""

    bold_font = Font(bold=True)
    for cell in worksheet[1]:
        cell.font = bold_font
        cell.fill = HEADER_FILL

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = max(len("" if cell.value is None else str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 40)


def apply_category_highlighting(worksheet: Worksheet, category_column_index: int) -> None:
    """Apply simple fill colours to usage categories."""

    if worksheet.max_row < 2:
        return

    column_letter = get_column_letter(category_column_index)
    data_range = f"{column_letter}2:{column_letter}{worksheet.max_row}"
    for category, fill in CATEGORY_FILLS.items():
        worksheet.conditional_formatting.add(
            data_range,
            CellIsRule(operator="equal", formula=[f'"{category}"'], fill=fill),
        )
