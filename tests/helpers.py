"""Сборка небольших книг и чтение итогового отчёта в тестах."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import REPORT_SHEET_NAME, REQUIRED_COLUMNS


def make_source(path: Path, sheets: list[tuple[str, list[str], list[dict]]]) -> None:
    """sheets: (название, заголовки, строки-словари)."""
    workbook = Workbook()
    for index, (title, headers, rows) in enumerate(sheets):
        worksheet = workbook.active if index == 0 else workbook.create_sheet(title)
        worksheet.title = title
        for column, header in enumerate(headers, start=1):
            worksheet.cell(1, column, header)
        for row_offset, row in enumerate(rows):
            for column, header in enumerate(headers, start=1):
                if header in row:
                    worksheet.cell(row_offset + 2, column, row[header])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def make_template(
    path: Path,
    headers: list[str] | None = None,
    *,
    data_font: str = "Calibri",
    quantity_format: str = "#,##0.00",
    header_fill: str = "1F4E79",
    column_width: float = 18,
) -> None:
    headers = list(headers or REQUIRED_COLUMNS)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = REPORT_SHEET_NAME
    fill = PatternFill(fill_type="solid", fgColor=header_fill)
    border = Border(
        left=Side(style="thin", color="D9E2F3"),
        right=Side(style="thin", color="D9E2F3"),
        top=Side(style="thin", color="D9E2F3"),
        bottom=Side(style="thin", color="D9E2F3"),
    )
    for column, header in enumerate(headers, start=1):
        header_cell = worksheet.cell(1, column, header)
        header_cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
        header_cell.fill = fill
        header_cell.alignment = Alignment(horizontal="center", vertical="center")
        prototype = worksheet.cell(2, column, None)
        prototype.font = Font(name=data_font, size=11, color="112233")
        prototype.border = border
        prototype.alignment = Alignment(vertical="center", horizontal="right")
        if header == "Date":
            prototype.number_format = "DD.MM.YYYY"
        elif header == "Quantity":
            prototype.number_format = quantity_format
        elif header in {"Unit Price", "Total"}:
            prototype.number_format = "#,##0.00"
        worksheet.column_dimensions[get_column_letter(column)].width = column_width
    worksheet.row_dimensions[1].height = 24
    worksheet.freeze_panes = "A2"
    worksheet.print_title_rows = "1:1"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def load_report(path: Path):
    workbook = load_workbook(path)
    worksheet = workbook[REPORT_SHEET_NAME]
    headers = []
    for column in range(1, worksheet.max_column + 1):
        value = worksheet.cell(1, column).value
        if value is not None:
            headers.append((column, value))
    rows = []
    totals = None
    for row_number in range(2, worksheet.max_row + 1):
        payload = {name: worksheet.cell(row_number, column).value for column, name in headers}
        if payload.get("Product") == "Итого" and not payload.get("SKU"):
            totals = payload
            totals["__row__"] = row_number
            continue
        if all(value is None for value in payload.values()):
            continue
        payload["__row__"] = row_number
        rows.append(payload)
    return workbook, worksheet, rows, totals


def as_date(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raise AssertionError(f"Ожидалась дата, получено {value!r}")


def assert_number(actual: object, expected: float, tolerance: float = 0.001) -> None:
    assert actual is not None, "число не записано"
    assert abs(float(actual) - expected) <= tolerance, f"{actual} != {expected}"


def rgb(color: object) -> str:
    if color is None or getattr(color, "rgb", None) is None:
        return ""
    return str(color.rgb).upper()
