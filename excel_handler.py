"""Чтение исходной книги и запись отчёта поверх копии шаблона."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.worksheet import Worksheet
from zipfile import BadZipFile

from config import HEADER_ROW, REPORT_SHEET_NAME, REQUIRED_COLUMNS
from processor import Issue, RawRow, ValidRecord


class ExcelAutomationError(Exception):
    """Ошибка, которую можно показать пользователю без трассировки."""


class SourceNotFoundError(ExcelAutomationError):
    pass


class CorruptWorkbookError(ExcelAutomationError):
    pass


class TemplateError(ExcelAutomationError):
    pass


class OutputWriteError(ExcelAutomationError):
    pass


class NoDataSheetsError(ExcelAutomationError):
    pass


_HEADER_INDEX = {name.casefold(): name for name in REQUIRED_COLUMNS}


@dataclass
class ReadResult:
    rows: list[RawRow] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    sheets_read: list[str] = field(default_factory=list)
    sheets_skipped: list[str] = field(default_factory=list)


def canonical_header(value: Any) -> str | None:
    """Возвращает стандартное имя колонки или None, если ячейка не заголовок."""
    if not isinstance(value, str):
        return None
    key = " ".join(value.strip().casefold().split())
    return _HEADER_INDEX.get(key)


def _ensure_xlsx(path: Path, label: str) -> None:
    if path.suffix.lower() != ".xlsx":
        raise ExcelAutomationError(
            f"{label} должен быть файлом .xlsx. Получен файл: {path.name}"
        )


def _load_workbook(path: Path, label: str) -> Any:
    try:
        return load_workbook(path)
    except PermissionError as exc:
        raise ExcelAutomationError(f"Нет доступа к файлу {label}: {path}") from exc
    except (InvalidFileException, BadZipFile, OSError, ValueError, KeyError) as exc:
        raise CorruptWorkbookError(
            f"Не удалось открыть {label}. Файл повреждён или это не книга Excel: {path}"
        ) from exc


def read_source(path: Path) -> ReadResult:
    """Читает все листы, на которых в первой строке есть нужные колонки.

    Лист без узнаваемых заголовков пропускается. Лист, где заголовки есть,
    но набор колонок неполный, пропускается и попадает в журнал как ошибка.
    """
    if not path.exists() or not path.is_file():
        raise SourceNotFoundError(f"Исходный файл не найден: {path}")
    _ensure_xlsx(path, "Исходный файл")

    workbook = _load_workbook(path, "исходный файл")
    result = ReadResult()
    try:
        for worksheet in workbook.worksheets:
            classified = _read_worksheet(worksheet)
            if classified is None:
                result.sheets_skipped.append(worksheet.title)
                result.issues.append(
                    Issue(
                        level="info",
                        sheet=worksheet.title,
                        row_number=None,
                        message="лист пропущен: в первой строке нет колонок с данными",
                    )
                )
                continue
            rows, missing = classified
            if missing:
                result.sheets_skipped.append(worksheet.title)
                result.issues.append(
                    Issue(
                        level="error",
                        sheet=worksheet.title,
                        row_number=None,
                        message="отсутствуют обязательные колонки: " + ", ".join(missing),
                    )
                )
                continue
            result.sheets_read.append(worksheet.title)
            result.rows.extend(rows)
    finally:
        workbook.close()
    return result


def _read_worksheet(worksheet: Worksheet) -> tuple[list[RawRow], list[str]] | None:
    found: dict[str, int] = {}
    for column in range(1, worksheet.max_column + 1):
        name = canonical_header(worksheet.cell(HEADER_ROW, column).value)
        if name is not None and name not in found:
            found[name] = column

    if not found:
        return None

    missing = [name for name in REQUIRED_COLUMNS if name not in found]
    rows: list[RawRow] = []
    if missing:
        return rows, missing

    for row_number in range(HEADER_ROW + 1, worksheet.max_row + 1):
        values = {
            name: worksheet.cell(row_number, column).value
            for name, column in found.items()
        }
        rows.append(RawRow(sheet=worksheet.title, row_number=row_number, values=values))
    return rows, []


def _capture_style(cell: Any) -> dict[str, Any]:
    return {
        "font": copy(cell.font),
        "fill": copy(cell.fill),
        "border": copy(cell.border),
        "alignment": copy(cell.alignment),
        "number_format": cell.number_format,
    }


def _apply_style(cell: Any, style: dict[str, Any]) -> None:
    cell.font = copy(style["font"])
    cell.fill = copy(style["fill"])
    cell.border = copy(style["border"])
    cell.alignment = copy(style["alignment"])
    cell.number_format = style["number_format"]


def _default_data_style(column_name: str) -> dict[str, Any]:
    number_formats = {
        "Date": "DD.MM.YYYY",
        "Quantity": "#,##0.00",
        "Unit Price": "#,##0.00",
        "Total": "#,##0.00",
    }
    horizontal = "right" if column_name in {"Quantity", "Unit Price", "Total"} else "left"
    thin = Border(
        left=Side(style="thin", color="D9E2F3"),
        right=Side(style="thin", color="D9E2F3"),
        top=Side(style="thin", color="D9E2F3"),
        bottom=Side(style="thin", color="D9E2F3"),
    )
    return {
        "font": Font(name="Calibri", size=11, color="1A1A1A"),
        "fill": PatternFill(fill_type=None),
        "border": thin,
        "alignment": Alignment(vertical="center", horizontal=horizontal),
        "number_format": number_formats.get(column_name, "General"),
    }


def write_report(records: list[ValidRecord], template_path: Path, output_path: Path) -> None:
    """Собирает итоговый файл из шаблона. Предыдущее содержимое output не дополняется.

    Оформление берётся из шаблона: стили заголовка не меняются, стиль строки данных
    копируется со второй строки шаблона. Лист справки остаётся как в шаблоне.
    """
    if not template_path.exists() or not template_path.is_file():
        raise TemplateError(f"Шаблон не найден: {template_path}")
    _ensure_xlsx(template_path, "Шаблон")
    _ensure_xlsx(output_path, "Итоговый файл")

    if template_path.resolve() == output_path.resolve():
        raise OutputWriteError("Нельзя записывать отчёт поверх файла шаблона.")
    if output_path.exists() and output_path.is_dir():
        raise OutputWriteError(f"Путь итогового файла указывает на папку: {output_path}")

    try:
        workbook = _load_workbook(template_path, "шаблон")
    except CorruptWorkbookError as exc:
        raise TemplateError(
            f"Шаблон повреждён или не является книгой Excel: {template_path}"
        ) from exc

    try:
        if REPORT_SHEET_NAME not in workbook.sheetnames:
            raise TemplateError(
                f"В шаблоне нет листа «{REPORT_SHEET_NAME}». "
                f"Доступные листы: {', '.join(workbook.sheetnames)}"
            )
        worksheet = workbook[REPORT_SHEET_NAME]
        columns = _template_columns(worksheet)
        styles = _data_styles(worksheet, columns)
        freeze_panes = worksheet.freeze_panes
        print_title_rows = worksheet.print_title_rows
        _clear_data_area(worksheet)
        _write_records(worksheet, records, columns, styles)
        _write_totals(worksheet, records, columns)
        _restore_sheet_view(
            worksheet,
            len(records),
            columns,
            freeze_panes=freeze_panes,
            print_title_rows=print_title_rows,
        )
        workbook.active = worksheet
        try:
            workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)
        except (TypeError, AttributeError):
            pass

        _save_workbook(workbook, output_path)
    finally:
        workbook.close()


def _template_columns(worksheet: Worksheet) -> dict[str, int]:
    found: dict[str, int] = {}
    for column in range(1, max(worksheet.max_column, len(REQUIRED_COLUMNS)) + 1):
        name = canonical_header(worksheet.cell(HEADER_ROW, column).value)
        if name is not None and name not in found:
            found[name] = column
    missing = [name for name in REQUIRED_COLUMNS if name not in found]
    if missing:
        raise TemplateError(
            "В шаблоне на листе «{sheet}» отсутствуют колонки: {columns}".format(
                sheet=REPORT_SHEET_NAME,
                columns=", ".join(missing),
            )
        )
    return found


def _data_styles(worksheet: Worksheet, columns: dict[str, int]) -> dict[str, dict[str, Any]]:
    """Берёт оформление из второй строки шаблона, если эта строка действительно есть.

    Обращение через worksheet.cell() создало бы пустые ячейки, поэтому смотрим
    только уже существующие.
    """
    styles: dict[str, dict[str, Any]] = {}
    existing = getattr(worksheet, "_cells", {})
    for name, column in columns.items():
        prototype = existing.get((2, column))
        if prototype is None:
            styles[name] = _default_data_style(name)
        else:
            styles[name] = _capture_style(prototype)
    return styles


def _clear_data_area(worksheet: Worksheet) -> None:
    """Удаляет строки ниже заголовка, чтобы повторная сборка не оставляла старые данные.

    Шаблон при каждом запуске открывается заново, поэтому удаление не затрагивает
    файл в папке templates.
    """
    if worksheet.max_row > HEADER_ROW:
        worksheet.delete_rows(HEADER_ROW + 1, worksheet.max_row - HEADER_ROW)


def _excel_number(value: Decimal) -> int | float:
    if value == value.to_integral_value():
        return int(value)
    return float(value)


def _record_value(record: ValidRecord, column_name: str) -> Any:
    if column_name == "Date":
        return record.date
    if column_name == "Product":
        return record.product
    if column_name == "SKU":
        return record.sku
    if column_name == "Quantity":
        return _excel_number(record.quantity)
    if column_name == "Unit Price":
        return _excel_number(record.unit_price)
    if column_name == "Total":
        return _excel_number(record.total)
    raise KeyError(column_name)


def _write_records(
    worksheet: Worksheet,
    records: list[ValidRecord],
    columns: dict[str, int],
    styles: dict[str, dict[str, Any]],
) -> None:
    for offset, record in enumerate(records):
        row_number = HEADER_ROW + 1 + offset
        for name, column in columns.items():
            cell = worksheet.cell(row_number, column, _record_value(record, name))
            _apply_style(cell, styles[name])


def _write_totals(
    worksheet: Worksheet,
    records: list[ValidRecord],
    columns: dict[str, int],
) -> None:
    if not records:
        return

    first_data = HEADER_ROW + 1
    last_data = HEADER_ROW + len(records)
    totals_row = last_data + 1
    label_column = columns["Product"]
    quantity_column = columns["Quantity"]
    total_column = columns["Total"]

    label = worksheet.cell(totals_row, label_column, "Итого")
    quantity = worksheet.cell(
        totals_row,
        quantity_column,
        f"={_sum_formula(quantity_column, first_data, last_data)}",
    )
    total = worksheet.cell(
        totals_row,
        total_column,
        f"={_sum_formula(total_column, first_data, last_data)}",
    )

    fill = PatternFill(fill_type="solid", fgColor="E7EEF7")
    font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
    border = Border(top=Side(style="medium", color="1F4E79"))
    for cell in (label, quantity, total):
        cell.fill = fill
        cell.font = font
        cell.border = border
    label.alignment = Alignment(vertical="center", horizontal="right")
    quantity.number_format = "#,##0.00"
    total.number_format = "#,##0.00"
    quantity.alignment = Alignment(vertical="center", horizontal="right")
    total.alignment = Alignment(vertical="center", horizontal="right")


def _sum_formula(column: int, first_row: int, last_row: int) -> str:
    letter = get_column_letter(column)
    return f"SUM({letter}{first_row}:{letter}{last_row})"


def _restore_sheet_view(
    worksheet: Worksheet,
    record_count: int,
    columns: dict[str, int],
    *,
    freeze_panes: str | None,
    print_title_rows: str | None,
) -> None:
    """Обновляет диапазон фильтра под число строк и возвращает закрепление из шаблона."""
    first_column = min(columns.values())
    last_column = max(columns.values())
    first_letter = get_column_letter(first_column)
    last_letter = get_column_letter(last_column)
    last_data = HEADER_ROW + record_count
    filter_end = max(last_data, HEADER_ROW)
    worksheet.auto_filter.ref = f"{first_letter}{HEADER_ROW}:{last_letter}{filter_end}"
    worksheet.freeze_panes = freeze_panes
    worksheet.print_title_rows = print_title_rows
    print_end = filter_end + (1 if record_count else 0)
    worksheet.print_area = f"{first_letter}{HEADER_ROW}:{last_letter}{print_end}"


def _save_workbook(workbook: Any, output_path: Path) -> None:
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise OutputWriteError(f"Не удалось создать папку для итогового файла: {output_path.parent}") from exc

    temporary = output_path.with_suffix(".xlsx.partial")
    try:
        workbook.save(temporary)
        temporary.replace(output_path)
    except OSError as exc:
        if temporary.exists():
            try:
                temporary.unlink()
            except OSError:
                pass
        raise OutputWriteError(f"Не удалось сохранить итоговый файл: {output_path}") from exc
