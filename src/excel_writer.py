"""Workbook writer and formatting for analysis outputs."""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from config import AppConfig
from src.models import AnalysisOutputs


HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9E2F3")
PRODUCTION_TECHNICIAN_FILL = PatternFill(fill_type="solid", fgColor="E2F0D9")
CATEGORY_FILLS = {
    "Regular": PatternFill(fill_type="solid", fgColor="C6EFCE"),
    "Occasional": PatternFill(fill_type="solid", fgColor="FFEB9C"),
    "Rare": PatternFill(fill_type="solid", fgColor="F4CCCC"),
}
REPORT_CHART_HEIGHT = 20
REPORT_CHART_WIDTH = 36
REPORT_SECTION_SPACER_ROWS = 42


def write_analysis_workbook(outputs: AnalysisOutputs, config: AppConfig) -> None:
    """Write all required sheets to a formatted Excel workbook."""

    output_path = config.output_file
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    workbook.remove(workbook.active)

    create_overview_sheet(workbook.create_sheet("Overview"), outputs)
    write_table_sheet(workbook.create_sheet("Raw_Data"), outputs.raw_data)
    write_table_sheet(
        workbook.create_sheet("User_Summary"),
        outputs.user_summary,
        category_column="usage_category",
        highlight_production_technicians=True,
    )
    write_table_sheet(
        workbook.create_sheet("Regular_Users"),
        outputs.regular_users,
        category_column="usage_category",
        highlight_production_technicians=True,
    )
    write_table_sheet(
        workbook.create_sheet("Occasional_Users"),
        outputs.occasional_users,
        category_column="usage_category",
        highlight_production_technicians=True,
    )
    write_table_sheet(
        workbook.create_sheet("Rare_Users"),
        outputs.rare_users,
        category_column="usage_category",
        highlight_production_technicians=True,
    )
    write_table_sheet(workbook.create_sheet("Monthly_Activity"), outputs.monthly_activity)
    write_table_sheet(workbook.create_sheet("Production_Techs"), outputs.production_technician_matches)
    write_table_sheet(workbook.create_sheet("Category_Rules"), outputs.category_rules)
    create_reporting_sheet(workbook.create_sheet("Reporting"), outputs)

    workbook.save(output_path)


def create_reporting_sheet(worksheet: Worksheet, outputs: AnalysisOutputs) -> None:
    """Create a reporting sheet with vertically stacked source tables and charts."""

    worksheet["A1"] = "PLM Usage Reporting"
    worksheet["A1"].font = Font(bold=True, size=16)

    current_row = 3
    current_row = write_reporting_section(worksheet, "Current Usage Facts", outputs.overview_metrics, current_row)

    category_header_row = current_row + 1
    current_row = write_reporting_section(
        worksheet,
        "Usage Category Split Source Data",
        outputs.category_summary,
        current_row,
    )
    add_category_split_chart(
        worksheet=worksheet,
        anchor_cell=f"A{current_row}",
        start_col=1,
        header_row=category_header_row,
        data_rows=len(outputs.category_summary),
    )
    current_row += REPORT_SECTION_SPACER_ROWS

    monthly_header_row = current_row + 1
    current_row = write_reporting_section(
        worksheet,
        "Monthly Active Users Source Data",
        outputs.monthly_active_users,
        current_row,
    )
    add_monthly_active_users_chart(
        worksheet=worksheet,
        anchor_cell=f"A{current_row}",
        start_col=1,
        header_row=monthly_header_row,
        data_rows=len(outputs.monthly_active_users),
    )
    current_row += REPORT_SECTION_SPACER_ROWS

    current_row = write_reporting_section(
        worksheet,
        "50 Most Active Users Source Data",
        outputs.most_active_users,
        current_row,
    )
    current_row += 4

    current_row = write_reporting_section(
        worksheet,
        "Users Not Logged In Last 90 Days",
        outputs.at_risk_rare_users,
        current_row,
    )
    current_row += 4

    current_row = write_reporting_section(
        worksheet,
        "Production Technician Match Review",
        outputs.production_technician_matches,
        current_row,
    )

    style_reporting_sheet(worksheet)


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
    highlight_production_technicians: bool = False,
) -> None:
    """Write a dataframe to a worksheet and apply basic formatting."""

    worksheet.append(list(dataframe.columns))
    for row in dataframe.itertuples(index=False, name=None):
        worksheet.append([excel_safe_value(value) for value in row])

    format_sheet(worksheet)
    if category_column and category_column in dataframe.columns:
        apply_category_highlighting(worksheet, dataframe.columns.get_loc(category_column) + 1)
    if highlight_production_technicians and "is_production_technician" in dataframe.columns:
        apply_production_technician_highlighting(
            worksheet,
            dataframe.columns.get_loc("is_production_technician") + 1,
        )


def write_reporting_section(
    worksheet: Worksheet,
    title: str,
    dataframe,
    start_row: int,
) -> int:
    """Write a reporting section vertically and return the next available row."""

    worksheet.cell(row=start_row, column=1, value=title)
    worksheet.cell(row=start_row, column=1).font = Font(bold=True, size=12)

    header_row = start_row + 1
    for column_index, header in enumerate(dataframe.columns, start=1):
        cell = worksheet.cell(row=header_row, column=column_index, value=header)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL

    for row_offset, row in enumerate(dataframe.itertuples(index=False, name=None), start=1):
        for column_index, value in enumerate(row, start=1):
            worksheet.cell(row=header_row + row_offset, column=column_index, value=excel_safe_value(value))

    return header_row + len(dataframe) + 3


def excel_safe_value(value):
    """Convert pandas missing values to blank Excel cells."""

    if pd.isna(value):
        return None
    return value


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


def style_reporting_sheet(worksheet: Worksheet) -> None:
    """Apply formatting specific to the reporting sheet."""

    worksheet.freeze_panes = "A4"
    worksheet.sheet_view.showGridLines = False

    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = max(len("" if cell.value is None else str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 14), 32)


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


def apply_production_technician_highlighting(
    worksheet: Worksheet,
    production_technician_column_index: int,
) -> None:
    """Highlight rows matched to the Production Technician list."""

    if worksheet.max_row < 2:
        return

    for row_index in range(2, worksheet.max_row + 1):
        if worksheet.cell(row=row_index, column=production_technician_column_index).value is True:
            for column_index in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row_index, column=column_index).fill = PRODUCTION_TECHNICIAN_FILL


def add_category_split_chart(
    worksheet: Worksheet,
    anchor_cell: str,
    start_col: int,
    header_row: int,
    data_rows: int,
) -> None:
    """Add a doughnut chart for user category split."""

    chart = DoughnutChart()
    chart.title = "Usage Category Split"
    chart.style = 10
    chart.height = REPORT_CHART_HEIGHT
    chart.width = REPORT_CHART_WIDTH

    labels = Reference(worksheet, min_col=start_col, min_row=header_row + 1, max_row=header_row + data_rows)
    data = Reference(worksheet, min_col=start_col + 1, min_row=header_row, max_row=header_row + data_rows)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showVal = True
    chart.dataLabels.showLeaderLines = True

    worksheet.add_chart(chart, anchor_cell)


def add_monthly_active_users_chart(
    worksheet: Worksheet,
    anchor_cell: str,
    start_col: int,
    header_row: int,
    data_rows: int,
) -> None:
    """Add a column chart for monthly active users."""

    chart = BarChart()
    chart.type = "col"
    chart.title = "Monthly Active Users by Month"
    chart.y_axis.title = "Active Users"
    chart.x_axis.title = "Reporting Month"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.style = 10
    chart.height = REPORT_CHART_HEIGHT
    chart.width = REPORT_CHART_WIDTH
    chart.legend = None
    chart.gapWidth = 35

    data = Reference(worksheet, min_col=start_col + 1, min_row=header_row, max_row=header_row + data_rows)
    categories = Reference(worksheet, min_col=start_col, min_row=header_row + 1, max_row=header_row + data_rows)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    worksheet.add_chart(chart, anchor_cell)


def add_ranked_users_chart(
    worksheet: Worksheet,
    anchor_cell: str,
    start_col: int,
    header_row: int,
    data_rows: int,
    title: str,
) -> None:
    """Add a horizontal bar chart for ranked user activity."""

    chart = BarChart()
    chart.type = "bar"
    chart.style = 12
    chart.title = title
    chart.y_axis.title = "User Name"
    chart.x_axis.title = "Average Active Days Per Month"
    chart.height = REPORT_CHART_HEIGHT
    chart.width = REPORT_CHART_WIDTH
    chart.legend = None
    chart.gapWidth = 40

    data = Reference(worksheet, min_col=start_col + 2, min_row=header_row, max_row=header_row + data_rows)
    categories = Reference(worksheet, min_col=start_col, min_row=header_row + 1, max_row=header_row + data_rows)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    worksheet.add_chart(chart, anchor_cell)
