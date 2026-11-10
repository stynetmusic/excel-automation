"""Создаёт демонстрационную книгу и шаблон отчёта.

Запуск из папки excel_automation:

    python sample_data.py

Существующие файлы input/source_sales.xlsx и templates/report_template.xlsx
будут перезаписаны. Итоговый отчёт в output эта команда не трогает.
"""

from __future__ import annotations

from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.worksheet import Worksheet

from config import (
    DEFAULT_SOURCE,
    DEFAULT_TEMPLATE,
    HELP_SHEET_NAME,
    REPORT_SHEET_NAME,
    REQUIRED_COLUMNS,
)

DATA_SHEETS = (
    "Store North",
    "Store South",
    "Store East",
    "Store West",
    "Online",
    "Wholesale",
)

CATALOG = (
    ("Coffee Beans 1kg", "CB-1001", Decimal("12.50")),
    ("Green Tea 500g", "GT-2002", Decimal("8.40")),
    ("Black Tea 500g", "BT-2003", Decimal("7.90")),
    ("Espresso Cups Set", "EC-3004", Decimal("24.00")),
    ("Milk Frother", "MF-3005", Decimal("35.50")),
    ("Paper Filters 100pcs", "PF-4006", Decimal("4.20")),
)

SHEET_COLORS = {
    "Store North": "1F4E79",
    "Store South": "0F6E56",
    "Store East": "7A4E2D",
    "Store West": "5C4B8A",
    "Online": "0E7490",
    "Wholesale": "9A3412",
}

HEADER_FILL = "1F4E79"
HEADER_FONT_COLOR = "FFFFFF"
HELP_MARKER = "Программа записывает данные только на лист «Сводный отчёт» и не изменяет этот лист."
COLUMN_WIDTHS = {"A": 16, "B": 32, "C": 18, "D": 14, "E": 16, "F": 16}


def _excel_value(value: object) -> object:
    if isinstance(value, Decimal):
        if value == value.to_integral_value():
            return int(value)
        return float(value)
    return value


def _append_bulk(sheets: dict[str, list[dict[str, object]]]) -> None:
    index = 0
    for day_offset in range(16):
        current = date(2026, 1, 5) + timedelta(days=day_offset)
        for product, sku, price in CATALOG:
            quantity = (index % 5) + 1
            total = (Decimal(quantity) * price).quantize(Decimal("0.01"))
            sheets[DATA_SHEETS[index % 6]].append(
                {
                    "date": current,
                    "product": product,
                    "sku": sku,
                    "quantity": quantity,
                    "price": price,
                    "total": total,
                }
            )
            index += 1
    if index != 96:
        raise RuntimeError(f"Ожидалось 96 массовых строк, получено {index}")


def _add(
    sheets: dict[str, list[dict[str, object]]],
    sheet: str,
    *,
    kind: str = "data",
    date_value: object = None,
    product: object = None,
    sku: object = None,
    quantity: object = None,
    price: object = None,
    total: object = None,
    comment: str = "",
) -> None:
    sheets[sheet].append(
        {
            "kind": kind,
            "date": date_value,
            "product": product,
            "sku": sku,
            "quantity": quantity,
            "price": price,
            "total": total,
            "comment": comment,
        }
    )


def _append_specials(sheets: dict[str, list[dict[str, object]]]) -> None:
    """Намеренно пустые, некорректные и повторные строки для проверки программы."""
    _add(sheets, "Store North", kind="empty", comment="Пустая строка: пропускается без ошибки.")
    _add(sheets, "Store North", kind="empty", comment="Вторая пустая строка.")
    _add(
        sheets,
        "Store North",
        kind="whitespace",
        comment="Строка из пробелов считается пустой.",
    )
    _add(
        sheets,
        "Store North",
        date_value=date(2026, 3, 1),
        product="Marker Calc",
        sku="CALC-001",
        quantity=4,
        price=Decimal("9.99"),
        total=None,
        comment="Total пустой: программа считает 4 × 9.99 = 39.96.",
    )
    _add(
        sheets,
        "Store North",
        date_value=date(2026, 2, 1),
        product="Marker Duplicate",
        sku="DUP-001",
        quantity=3,
        price=Decimal("12.50"),
        total=Decimal("37.50"),
        comment="Первое вхождение дубликата. В отчёте остаётся количество 3.",
    )
    _add(
        sheets,
        "Store North",
        date_value=date(2026, 3, 2),
        product="Marker Negative Qty",
        sku="NEG-001",
        quantity=-2,
        price=Decimal("8.40"),
        total=None,
        comment="Отрицательное количество: строка отклоняется.",
    )
    _add(
        sheets,
        "Store North",
        date_value="вчера",
        product="Marker Bad Date",
        sku="BADDATE-001",
        quantity=1,
        price=Decimal("7.90"),
        total=Decimal("7.90"),
        comment="Некорректная дата: строка отклоняется.",
    )

    _add(
        sheets,
        "Store South",
        date_value=date(2026, 3, 3),
        product="Marker Match",
        sku="MATCH-001",
        quantity=2,
        price=Decimal("24.00"),
        total=Decimal("48.00"),
        comment="Total совпадает с расчётом.",
    )
    _add(
        sheets,
        "Store South",
        date_value=date(2026, 3, 4),
        product="Marker Tolerance",
        sku="TOL-001",
        quantity=1,
        price=Decimal("35.50"),
        total="35.51",
        comment="Расхождение Total ровно 0.01: строка принимается.",
    )
    _add(
        sheets,
        "Store South",
        date_value=date(2026, 3, 5),
        product="Marker Bad Total",
        sku="BADTOTAL-001",
        quantity=2,
        price=Decimal("4.20"),
        total=Decimal("99"),
        comment="Total 99 не равен 8.40: строка отклоняется.",
    )
    _add(
        sheets,
        "Store South",
        date_value=date(2026, 3, 6),
        product=None,
        sku="NOPROD-001",
        quantity=1,
        price=Decimal("9.99"),
        total=Decimal("9.99"),
        comment="Пустой Product: строка отклоняется.",
    )
    _add(
        sheets,
        "Store South",
        date_value=date(2026, 2, 1),
        product="Marker Duplicate",
        sku="DUP-001",
        quantity=9,
        price=Decimal("12.50"),
        total=Decimal("112.50"),
        comment="Повтор Date+SKU+Product с другого листа: дубликат.",
    )

    _add(
        sheets,
        "Store East",
        date_value=date(2026, 3, 7),
        product="Marker No SKU",
        sku=None,
        quantity=1,
        price=Decimal("8.40"),
        total=Decimal("8.40"),
        comment="Пустой SKU: строка отклоняется.",
    )
    _add(
        sheets,
        "Store East",
        date_value=date(2026, 3, 8),
        product="Marker No Qty",
        sku="NOQTY-001",
        quantity=None,
        price=Decimal("7.90"),
        total=None,
        comment="Пустое количество: строка отклоняется.",
    )
    _add(
        sheets,
        "Store East",
        date_value=date(2026, 3, 9),
        product="Marker Negative Price",
        sku="NEGPRICE-001",
        quantity=1,
        price=Decimal("-5"),
        total=None,
        comment="Отрицательная цена: строка отклоняется.",
    )
    _add(
        sheets,
        "Store East",
        date_value=date(2026, 3, 10),
        product="Marker Bad Price",
        sku="BADPRICE-001",
        quantity=1,
        price="N/A",
        total=None,
        comment="Цена текстом: строка отклоняется.",
    )
    _add(
        sheets,
        "Store East",
        date_value=date(2026, 3, 11),
        product="Marker Bad Qty",
        sku="BADQTY-001",
        quantity="много",
        price=Decimal("6.30"),
        total=None,
        comment="Количество текстом: строка отклоняется.",
    )
    _add(
        sheets,
        "Store East",
        date_value=date(2026, 2, 10),
        product="Marker Sheet Dup",
        sku="DUP-SHEET",
        quantity=2,
        price=Decimal("6.30"),
        total=Decimal("12.60"),
        comment="Первая из двух одинаковых строк на этом листе. Остаётся количество 2.",
    )
    _add(
        sheets,
        "Store East",
        date_value=date(2026, 2, 10),
        product="Marker Sheet Dup",
        sku="DUP-SHEET",
        quantity=5,
        price=Decimal("6.30"),
        total=Decimal("31.50"),
        comment="Повтор на том же листе: дубликат.",
    )

    _add(
        sheets,
        "Store West",
        date_value="15.03.2026",
        product="Marker Date",
        sku="DATE-001",
        quantity=3,
        price=Decimal("5.60"),
        total=None,
        comment="Дата текстом ДД.ММ.ГГГГ. Total считается: 16.80.",
    )
    _add(
        sheets,
        "Store West",
        date_value="2026-03-16",
        product="Marker ISO",
        sku="ISO-001",
        quantity=1,
        price=Decimal("18.75"),
        total=Decimal("18.75"),
        comment="Дата текстом ГГГГ-ММ-ДД.",
    )
    _add(
        sheets,
        "Store West",
        date_value=date(2026, 3, 17),
        product="Marker Zero",
        sku="ZERO-001",
        quantity=0,
        price=Decimal("9.99"),
        total=0,
        comment="Нулевое количество допустимо. Отрицательное — нет.",
    )
    _add(
        sheets,
        "Store West",
        date_value=None,
        product="Marker No Date",
        sku="NODATE-001",
        quantity=1,
        price=Decimal("4.20"),
        total=Decimal("4.20"),
        comment="Пустая дата: строка отклоняется.",
    )
    _add(
        sheets,
        "Store West",
        date_value=date(2026, 3, 26),
        product="Marker No Price",
        sku="NOPRICE-001",
        quantity=1,
        price=None,
        total=None,
        comment="Пустая цена: строка отклоняется.",
    )
    _add(
        sheets,
        "Store West",
        date_value=date(2026, 3, 12),
        product="Marker Off Total",
        sku="OFFTOTAL-001",
        quantity=1,
        price=Decimal("10.00"),
        total=Decimal("10.02"),
        comment="Расхождение 0.02 больше допуска 0.01: строка отклоняется.",
    )

    _add(
        sheets,
        "Online",
        date_value=date(2026, 3, 18),
        product="  Marker Space  ",
        sku="  SPACE-001  ",
        quantity=2,
        price=Decimal("9.99"),
        total=None,
        comment="Лишние пробелы по краям снимаются. Total = 19.98.",
    )
    _add(
        sheets,
        "Online",
        date_value=date(2026, 3, 19),
        product="Marker Case",
        sku="CASE-001",
        quantity=1,
        price=Decimal("9.99"),
        total=Decimal("9.99"),
        comment="Регистр входит в ключ. Эта строка и следующая — не дубликаты.",
    )
    _add(
        sheets,
        "Online",
        date_value=date(2026, 3, 19),
        product="marker case",
        sku="CASE-001",
        quantity=1,
        price=Decimal("9.99"),
        total=Decimal("9.99"),
        comment="Тот же SKU и дата, но другое написание Product.",
    )
    _add(
        sheets,
        "Online",
        date_value=date(2026, 3, 20),
        product="Marker String Total",
        sku="STRTOT-001",
        quantity=3,
        price=Decimal("12.50"),
        total="37.50",
        comment="Total записан текстом и совпадает с расчётом.",
    )
    _add(
        sheets,
        "Online",
        date_value=date(2026, 3, 25),
        product="Marker Euro Price",
        sku="EUR-001",
        quantity=2,
        price="8,40",
        total=None,
        comment="Цена с десятичной запятой. Total = 16.80.",
    )
    _add(
        sheets,
        "Online",
        date_value=date(2026, 2, 1),
        product="Marker Duplicate",
        sku="DUP-001",
        quantity=1,
        price=Decimal("12.50"),
        total=Decimal("12.50"),
        comment="Третье вхождение DUP-001: тоже дубликат.",
    )

    _add(
        sheets,
        "Wholesale",
        date_value=date(2026, 3, 21),
        product="Marker Float Qty",
        sku="QTYFLOAT-001",
        quantity=Decimal("2.5"),
        price=Decimal("4.20"),
        total=Decimal("10.50"),
        comment="Дробное количество 2.5.",
    )
    _add(
        sheets,
        "Wholesale",
        date_value=date(2026, 3, 22),
        product="Marker Free",
        sku="FREE-001",
        quantity=1,
        price=0,
        total=0,
        comment="Нулевая цена допустима.",
    )
    _add(
        sheets,
        "Wholesale",
        date_value=date(2026, 3, 23),
        product="Marker Tolerance 2",
        sku="TOL-002",
        quantity=3,
        price=Decimal("1.00"),
        total="3.009",
        comment="Расхождение 0.009 в пределах допуска. В отчёт попадёт 3.01.",
    )
    _add(
        sheets,
        "Wholesale",
        date_value=date(2026, 3, 24),
        product="Marker Text Qty",
        sku="STRQTY-001",
        quantity="4",
        price="9.99",
        total=None,
        comment="Количество и цена записаны текстом. Total = 39.96.",
    )


def _style_source_header(cell: object, color: str) -> None:
    cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill(fill_type="solid", fgColor=color)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _write_source_sheet(workbook: Workbook, title: str, rows: list[dict[str, object]], first: bool) -> None:
    worksheet: Worksheet = workbook.active if first else workbook.create_sheet(title)
    worksheet.title = title
    color = SHEET_COLORS[title]
    worksheet.sheet_properties.tabColor = color
    headers = list(REQUIRED_COLUMNS) + ["Comment"]
    for column, header in enumerate(headers, start=1):
        cell = worksheet.cell(1, column, header)
        _style_source_header(cell, color)
    worksheet.row_dimensions[1].height = 22
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:G{max(1, len(rows) + 1)}"
    widths = {"A": 16, "B": 28, "C": 18, "D": 14, "E": 14, "F": 14, "G": 78}
    for letter, width in widths.items():
        worksheet.column_dimensions[letter].width = width

    for offset, item in enumerate(rows):
        row_number = offset + 2
        kind = item.get("kind", "data")
        if kind == "empty":
            values = [None, None, None, None, None, None]
        elif kind == "whitespace":
            values = [" ", " ", " ", " ", " ", " "]
        else:
            values = [
                item.get("date"),
                item.get("product"),
                item.get("sku"),
                _excel_value(item.get("quantity")),
                _excel_value(item.get("price")),
                _excel_value(item.get("total")),
            ]
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row_number, column, value)
            if column == 1 and isinstance(value, date):
                cell.number_format = "DD.MM.YYYY"
            elif column in {4, 5, 6} and isinstance(value, (int, float)):
                cell.number_format = "#,##0.00"
        comment = item.get("comment")
        if comment:
            worksheet.cell(row_number, 7, comment)


def build_source(path: Path) -> None:
    sheets = {name: [] for name in DATA_SHEETS}
    _append_bulk(sheets)
    _append_specials(sheets)

    workbook = Workbook()
    for index, name in enumerate(DATA_SHEETS):
        _write_source_sheet(workbook, name, sheets[name], first=index == 0)

    draft = workbook.create_sheet("Черновик")
    draft.sheet_properties.tabColor = "B91C1C"
    for column, header in enumerate(("Date", "Product", "SKU"), start=1):
        cell = draft.cell(1, column, header)
        _style_source_header(cell, "B91C1C")
    draft.cell(2, 1, date(2026, 3, 1))
    draft.cell(2, 2, "Не должен попасть в отчёт")
    draft.cell(2, 3, "DRAFT-001")
    draft.cell(1, 4, "Нет колонок Quantity, Unit Price и Total. Весь лист пропускается.")
    draft.column_dimensions["D"].width = 78

    notes = workbook.create_sheet("Инструкция")
    notes.sheet_properties.tabColor = "64748B"
    notes["A1"] = "Служебный лист"
    notes["A1"].font = Font(name="Calibri", bold=True, size=14, color="1F4E79")
    notes["A3"] = (
        "Это не таблица продаж. Здесь нет строки заголовков Date, Product, SKU, "
        "Quantity, Unit Price и Total, поэтому программа пропускает лист."
    )
    notes["A4"] = (
        "Колонка Comment на листах магазинов — пояснение к примеру. "
        "В итоговый отчёт она не переносится."
    )
    notes.column_dimensions["A"].width = 110

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _thin_border(color: str) -> Border:
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def build_template(path: Path) -> None:
    workbook = Workbook()
    report = workbook.active
    report.title = REPORT_SHEET_NAME
    report.sheet_properties.tabColor = HEADER_FILL

    header_font = Font(name="Calibri", bold=True, color=HEADER_FONT_COLOR, size=12)
    header_fill = PatternFill(fill_type="solid", fgColor=HEADER_FILL)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_border = _thin_border("16385C")

    data_font = Font(name="Calibri", size=11, color="1A1A1A")
    data_border = _thin_border("D9E2F3")
    number_formats = {
        "Date": "DD.MM.YYYY",
        "Product": "General",
        "SKU": "General",
        "Quantity": "#,##0.00",
        "Unit Price": "#,##0.00",
        "Total": "#,##0.00",
    }
    alignments = {
        "Date": Alignment(vertical="center", horizontal="center"),
        "Product": Alignment(vertical="center", horizontal="left"),
        "SKU": Alignment(vertical="center", horizontal="left"),
        "Quantity": Alignment(vertical="center", horizontal="right"),
        "Unit Price": Alignment(vertical="center", horizontal="right"),
        "Total": Alignment(vertical="center", horizontal="right"),
    }

    for column, name in enumerate(REQUIRED_COLUMNS, start=1):
        header = report.cell(1, column, name)
        header.font = header_font
        header.fill = header_fill
        header.alignment = header_alignment
        header.border = header_border

        prototype = report.cell(2, column, None)
        prototype.font = data_font
        prototype.border = data_border
        prototype.alignment = alignments[name]
        prototype.number_format = number_formats[name]
        report.column_dimensions[get_column_letter(column)].width = COLUMN_WIDTHS[get_column_letter(column)]

    report.row_dimensions[1].height = 24
    report.row_dimensions[2].height = 18
    report.freeze_panes = "A2"
    report.auto_filter.ref = "A1:F2"
    report.print_title_rows = "1:1"
    report.page_setup.orientation = "landscape"
    report.page_setup.paperSize = report.PAPERSIZE_A4
    report.page_setup.fitToWidth = 1
    report.page_setup.fitToHeight = 0
    report.page_setup.horizontalCentered = True
    report.sheet_properties.pageSetUpPr.fitToPage = True
    report.page_setup.horizontalCentered = True
    report.oddHeader.left.text = "Сводный отчёт по продажам"
    report.oddFooter.right.text = "Стр. &P из &N"
    report.sheet_view.showGridLines = False
    report.page_setup.scale = 100
    report.sheet_view.zoomScale = 120

    help_sheet = workbook.create_sheet(HELP_SHEET_NAME)
    help_sheet.sheet_properties.tabColor = "C65911"
    help_sheet["A1"] = "Справка по шаблону"
    help_sheet["A1"].font = Font(name="Calibri", size=16, bold=True, color=HEADER_FILL)
    help_sheet["A1"].fill = PatternFill(fill_type="solid", fgColor="F4E1D2")
    help_sheet["A3"] = HELP_MARKER
    help_sheet["A4"] = (
        "Вторая строка листа «Сводный отчёт» задаёт шрифт, границы и формат чисел "
        "для будущих строк. При запуске программы она заменяется данными."
    )
    help_sheet["A5"] = (
        "Сам файл шаблона не изменяется. Каждый запуск заново создаёт файл в папке output."
    )
    help_sheet["A6"] = (
        "Можно менять цвета, ширину колонок и шрифт. Не переименовывайте колонки "
        "и не удаляйте лист «Сводный отчёт»."
    )
    help_sheet.column_dimensions["A"].width = 110
    help_sheet.row_dimensions[1].height = 26
    help_sheet.sheet_view.showGridLines = False

    workbook.properties.title = "Сводный отчёт по продажам"
    workbook.properties.creator = "Excel Automation"
    workbook.properties.subject = "Шаблон консолидации продаж"
    workbook.active = report
    try:
        workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)
    except (TypeError, AttributeError):
        pass

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def main() -> None:
    build_source(DEFAULT_SOURCE)
    build_template(DEFAULT_TEMPLATE)
    print(f"Исходная книга: {DEFAULT_SOURCE}")
    print(f"Шаблон:         {DEFAULT_TEMPLATE}")

    from excel_handler import read_source
    from processor import process_rows

    read_result = read_source(DEFAULT_SOURCE)
    processed = process_rows(read_result.rows)
    sheet_errors = sum(1 for issue in read_result.issues if issue.level == "error")
    print(f"Листов с данными: {len(read_result.sheets_read)}")
    print(f"Прочитано строк: {processed.rows_read}")
    print(f"Успешно обработано: {processed.processed}")
    print(f"Пропущено записей: {processed.skipped}")
    print(f"Обнаружено дубликатов: {processed.duplicates}")
    print(f"Пустых строк: {processed.empty_rows}")
    print(f"Листов с ошибкой структуры: {sheet_errors}")


if __name__ == "__main__":
    main()
