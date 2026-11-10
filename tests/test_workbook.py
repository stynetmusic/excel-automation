"""Проверки, которые читают реальные ячейки Excel, а не только возвращаемые объекты."""

from __future__ import annotations

import hashlib
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from config import REPORT_SHEET_NAME, REQUIRED_COLUMNS
from excel_handler import read_source, write_report
from main import main
from processor import process_rows
from tests.helpers import (
    as_date,
    assert_number,
    load_report,
    make_source,
    make_template,
    rgb,
)

ROOT = Path(__file__).resolve().parents[1]
SAMPLE_SOURCE = ROOT / "input" / "source_sales.xlsx"
SAMPLE_TEMPLATE = ROOT / "templates" / "report_template.xlsx"


def _sale(sku: str, product: str = "Товар", **extra: object) -> dict[str, object]:
    row: dict[str, object] = {
        "Date": date(2026, 5, 1),
        "Product": product,
        "SKU": sku,
        "Quantity": 2,
        "Unit Price": 10,
        "Total": 20,
    }
    row.update(extra)
    return row


def test_reads_six_sheets_and_skips_incomplete_ones(tmp_path: Path) -> None:
    sheets = []
    for index, name in enumerate(("North", "South", "East", "West", "Online", "Wholesale")):
        headers = ["Date", "Region", "Product", "SKU", "Quantity", "Unit Price", "Total"]
        row = _sale(f"S{index}", product=f"Позиция {index}", Region="Север")
        sheets.append((name, headers, [row]))
    sheets.append(("Инструкция", ["Заметка"], [{"Заметка": "не таблица"}]))
    sheets.append(
        (
            "Черновик",
            ["Date", "Product", "SKU"],
            [{"Date": date(2026, 5, 2), "Product": "Черновик", "SKU": "DRAFT-X"}],
        )
    )
    source = tmp_path / "six.xlsx"
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    make_source(source, sheets)
    make_template(template)

    read_result = read_source(source)
    assert read_result.sheets_read == ["North", "South", "East", "West", "Online", "Wholesale"]
    assert any(issue.sheet == "Черновик" and "обязательные колонки" in issue.message for issue in read_result.issues)

    code = main(
        ["--input", str(source), "--template", str(template), "--output", str(output), "--log", str(tmp_path / "app.log")]
    )
    assert code == 0
    _workbook, worksheet, rows, _totals = load_report(output)
    assert [row["SKU"] for row in rows] == [f"S{index}" for index in range(6)]
    assert [row["Product"] for row in rows] == [f"Позиция {index}" for index in range(6)]
    assert "Region" not in [cell.value for cell in worksheet[1]]
    assert all(row["SKU"] != "DRAFT-X" for row in rows)
    _workbook.close()


def test_merge_total_empty_rows_and_duplicates_in_cells(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    make_source(
        source,
        [
            (
                "Store North",
                list(REQUIRED_COLUMNS),
                [
                    _sale("KEEP", Quantity=4, **{"Unit Price": 9.99, "Total": None}),
                    {"Date": None, "Product": None, "SKU": None, "Quantity": None, "Unit Price": None, "Total": None},
                    {"Date": " ", "Product": " ", "SKU": " ", "Quantity": " ", "Unit Price": " ", "Total": " "},
                    _sale("NEG", Quantity=-1, Total=None),
                    _sale("DUP", product="Один", Quantity=3, **{"Unit Price": 12.5, "Total": 37.5}),
                ],
            ),
            (
                "Store South",
                list(REQUIRED_COLUMNS),
                [
                    _sale("MATCH", Quantity=2, **{"Unit Price": 24, "Total": 48}),
                    _sale("TOL", Quantity=1, **{"Unit Price": 35.5, "Total": 35.51}),
                    _sale("BAD", Quantity=1, **{"Unit Price": 10, "Total": 10.02}),
                    _sale("DUP", product="Один", Quantity=9, **{"Unit Price": 12.5, "Total": 112.5}),
                ],
            ),
        ],
    )
    make_template(template)
    code = main(
        ["--input", str(source), "--template", str(template), "--output", str(output), "--log", str(tmp_path / "app.log")]
    )
    assert code == 0
    workbook, _worksheet, rows, totals = load_report(output)
    by_sku = {row["SKU"]: row for row in rows}
    assert set(by_sku) == {"KEEP", "DUP", "MATCH", "TOL"}
    assert_number(by_sku["KEEP"]["Total"], 39.96)
    assert_number(by_sku["KEEP"]["Quantity"], 4)
    assert_number(by_sku["MATCH"]["Total"], 48)
    assert_number(by_sku["TOL"]["Total"], 35.51)
    assert_number(by_sku["DUP"]["Quantity"], 3)
    assert "BAD" not in by_sku
    assert "NEG" not in by_sku
    assert totals is not None
    assert str(totals["Quantity"]).startswith("=SUM(")
    assert str(totals["Total"]).startswith("=SUM(")
    log_text = (tmp_path / "app.log").read_text(encoding="utf-8")
    assert "отрицательное количество" in log_text
    assert "дубликат" in log_text
    assert "не совпадает" in log_text
    assert "Traceback" not in log_text
    workbook.close()


def test_template_styles_are_copied_onto_data_cells(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    make_source(source, [("Лист", list(REQUIRED_COLUMNS), [_sale("STYLED")])])
    make_template(
        template,
        data_font="Georgia",
        quantity_format="0.000",
        header_fill="112233",
        column_width=21,
    )
    before = hashlib.sha256(template.read_bytes()).hexdigest()
    write_rows = read_source(source)
    result = process_rows(write_rows.rows)
    write_report(result.records, template, output)
    assert hashlib.sha256(template.read_bytes()).hexdigest() == before

    workbook, worksheet, rows, _totals = load_report(output)
    assert len(rows) == 1
    header = worksheet.cell(1, 1)
    assert header.value == "Date"
    assert header.font.bold is True
    assert header.font.color is not None and rgb(header.font.color).endswith("FFFFFF")
    assert rgb(header.fill.fgColor).endswith("112233")
    assert worksheet.column_dimensions["A"].width == 21
    assert worksheet.freeze_panes == "A2"
    assert worksheet.print_title_rows == "$1:$1"

    quantity_column = next(column for column, name in enumerate(REQUIRED_COLUMNS, start=1) if name == "Quantity")
    data_cell = worksheet.cell(2, quantity_column)
    assert data_cell.font.name == "Georgia"
    assert data_cell.font.color is not None and rgb(data_cell.font.color).endswith("112233")
    assert data_cell.number_format == "0.000"
    assert header.font.name == "Calibri"
    workbook.close()


def test_output_follows_template_column_order(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    reversed_headers = ["Total", "Unit Price", "Quantity", "SKU", "Product", "Date"]
    make_source(source, [("Лист", list(REQUIRED_COLUMNS), [_sale("ORDER", product="Порядок")])])
    make_template(template, headers=reversed_headers)
    assert main(
        ["--input", str(source), "--template", str(template), "--output", str(output), "--log", str(tmp_path / "app.log")]
    ) == 0
    workbook, worksheet, rows, totals = load_report(output)
    assert [worksheet.cell(1, column).value for column in range(1, 7)] == reversed_headers
    assert rows[0]["SKU"] == "ORDER"
    assert rows[0]["Product"] == "Порядок"
    assert as_date(rows[0]["Date"]) == date(2026, 5, 1)
    sku_column = 4
    assert worksheet.cell(rows[0]["__row__"], sku_column).value == "ORDER"
    assert totals is not None
    total_letter = get_column_letter(1)
    quantity_letter = get_column_letter(3)
    assert totals["Total"] == f"=SUM({total_letter}2:{total_letter}2)"
    assert totals["Quantity"] == f"=SUM({quantity_letter}2:{quantity_letter}2)"
    workbook.close()


def test_rerun_rebuilds_file_instead_of_appending(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    log_path = tmp_path / "app.log"
    make_template(template, header_fill="445566")
    make_source(
        source,
        [("Лист", list(REQUIRED_COLUMNS), [_sale("ONE", product="Первый"), _sale("TWO", product="Второй", Date=date(2026, 5, 2))])],
    )
    args = ["--input", str(source), "--template", str(template), "--output", str(output), "--log", str(log_path)]
    assert main(args) == 0

    workbook = load_workbook(output)
    worksheet = workbook[REPORT_SHEET_NAME]
    worksheet.cell(2, 3, "GHOST")
    workbook.save(output)
    workbook.close()

    make_source(source, [("Лист", list(REQUIRED_COLUMNS), [_sale("ONE", product="Первый")])])
    assert main(args) == 0
    workbook, worksheet, rows, _totals = load_report(output)
    assert [row["SKU"] for row in rows] == ["ONE"]
    assert all(row["SKU"] != "GHOST" for row in rows)
    assert rgb(worksheet.cell(1, 1).fill.fgColor).endswith("445566")
    values_first = [(row["SKU"], row["Product"]) for row in rows]
    assert main(args) == 0
    _workbook, _worksheet, rows_again, _totals = load_report(output)
    assert [(row["SKU"], row["Product"]) for row in rows_again] == values_first
    workbook.close()
    _workbook.close()


def test_sample_workbook_content_and_formatting(tmp_path: Path) -> None:
    assert SAMPLE_SOURCE.exists()
    assert SAMPLE_TEMPLATE.exists()
    output = tmp_path / "consolidated.xlsx"
    log_path = tmp_path / "app.log"
    template_hash = hashlib.sha256(SAMPLE_TEMPLATE.read_bytes()).hexdigest()
    source_hash = hashlib.sha256(SAMPLE_SOURCE.read_bytes()).hexdigest()

    code = main(
        [
            "--input",
            str(SAMPLE_SOURCE),
            "--template",
            str(SAMPLE_TEMPLATE),
            "--output",
            str(output),
            "--log",
            str(log_path),
        ]
    )
    assert code == 0
    assert hashlib.sha256(SAMPLE_TEMPLATE.read_bytes()).hexdigest() == template_hash
    assert hashlib.sha256(SAMPLE_SOURCE.read_bytes()).hexdigest() == source_hash

    raw_book = load_workbook(SAMPLE_SOURCE)
    nonempty = 0
    data_sheets = []
    for worksheet in raw_book.worksheets:
        headers = [worksheet.cell(1, column).value for column in range(1, worksheet.max_column + 1)]
        if not all(name in headers for name in REQUIRED_COLUMNS):
            continue
        data_sheets.append(worksheet.title)
        columns = {name: headers.index(name) + 1 for name in REQUIRED_COLUMNS}
        for row_number in range(2, worksheet.max_row + 1):
            values = [worksheet.cell(row_number, columns[name]).value for name in REQUIRED_COLUMNS]
            if any(value is not None and not (isinstance(value, str) and value.strip() == "") for value in values):
                nonempty += 1
    raw_book.close()
    assert data_sheets == [
        "Store North",
        "Store South",
        "Store East",
        "Store West",
        "Online",
        "Wholesale",
    ]
    assert nonempty >= 100
    assert nonempty == 128

    workbook, worksheet, rows, totals = load_report(output)
    assert workbook.sheetnames == ["Сводный отчёт", "Справка"]
    assert len(rows) == 113
    assert worksheet.freeze_panes == "A2"
    assert worksheet.column_dimensions["A"].width == 16
    assert worksheet.column_dimensions["B"].width == 32
    assert worksheet.print_title_rows == "$1:$1"
    assert worksheet.page_setup.orientation == "landscape"
    assert worksheet.sheet_view.showGridLines is False
    assert worksheet.oddHeader.left.text == "Сводный отчёт по продажам"
    header = worksheet.cell(1, 1)
    assert header.font.bold is True
    assert rgb(header.fill.fgColor).endswith("1F4E79")
    assert header.font.color is not None and rgb(header.font.color).endswith("FFFFFF")
    assert worksheet.cell(2, 1).number_format == "DD.MM.YYYY"
    assert worksheet.cell(2, 6).number_format == "#,##0.00"
    help_sheet = workbook["Справка"]
    assert help_sheet["A1"].value == "Справка по шаблону"
    assert str(help_sheet["A3"].value).startswith("Программа записывает данные только на лист")
    assert rgb(help_sheet["A1"].fill.fgColor).endswith("F4E1D2")

    by_key = {(row["SKU"], row["Product"]): row for row in rows}
    calc = by_key[("CALC-001", "Marker Calc")]
    assert as_date(calc["Date"]) == date(2026, 3, 1)
    assert_number(calc["Quantity"], 4)
    assert_number(calc["Unit Price"], 9.99)
    assert_number(calc["Total"], 39.96)

    assert_number(by_key[("MATCH-001", "Marker Match")]["Total"], 48)
    assert_number(by_key[("TOL-001", "Marker Tolerance")]["Total"], 35.51)
    assert_number(by_key[("DUP-001", "Marker Duplicate")]["Quantity"], 3)
    assert_number(by_key[("DUP-SHEET", "Marker Sheet Dup")]["Quantity"], 2)
    assert as_date(by_key[("DATE-001", "Marker Date")]["Date"]) == date(2026, 3, 15)
    assert_number(by_key[("DATE-001", "Marker Date")]["Total"], 16.80)
    assert as_date(by_key[("ISO-001", "Marker ISO")]["Date"]) == date(2026, 3, 16)
    assert_number(by_key[("ZERO-001", "Marker Zero")]["Quantity"], 0)
    assert by_key[("SPACE-001", "Marker Space")]["SKU"] == "SPACE-001"
    assert_number(by_key[("SPACE-001", "Marker Space")]["Total"], 19.98)
    assert ("CASE-001", "Marker Case") in by_key
    assert ("CASE-001", "marker case") in by_key
    assert_number(by_key[("STRTOT-001", "Marker String Total")]["Total"], 37.50)
    assert_number(by_key[("EUR-001", "Marker Euro Price")]["Total"], 16.80)
    assert_number(by_key[("EUR-001", "Marker Euro Price")]["Unit Price"], 8.40)
    assert_number(by_key[("QTYFLOAT-001", "Marker Float Qty")]["Quantity"], 2.5)
    assert_number(by_key[("QTYFLOAT-001", "Marker Float Qty")]["Total"], 10.50)
    assert_number(by_key[("FREE-001", "Marker Free")]["Unit Price"], 0)
    assert_number(by_key[("TOL-002", "Marker Tolerance 2")]["Total"], 3.01)
    assert_number(by_key[("STRQTY-001", "Marker Text Qty")]["Total"], 39.96)

    coffee = [
        row
        for row in rows
        if row["SKU"] == "CB-1001" and as_date(row["Date"]) == date(2026, 1, 5)
    ]
    assert len(coffee) == 1
    assert coffee[0]["Product"] == "Coffee Beans 1kg"
    assert_number(coffee[0]["Quantity"], 1)
    assert_number(coffee[0]["Total"], 12.50)
    assert sum(1 for row in rows if row["SKU"] == "CB-1001") == 16

    rejected = {
        "NEG-001",
        "BADDATE-001",
        "BADTOTAL-001",
        "NOPROD-001",
        "NOQTY-001",
        "NEGPRICE-001",
        "BADPRICE-001",
        "BADQTY-001",
        "NODATE-001",
        "NOPRICE-001",
        "OFFTOTAL-001",
        "DRAFT-001",
    }
    present = {row["SKU"] for row in rows}
    assert rejected.isdisjoint(present)
    assert all(row["Product"] != "Marker No SKU" for row in rows)
    assert all(row["Product"] != "Не должен попасть в отчёт" for row in rows)

    keys = [(as_date(row["Date"]), row["SKU"], row["Product"]) for row in rows]
    assert len(keys) == len(set(keys))

    assert totals is not None
    assert totals["__row__"] == 115
    assert totals["Quantity"] == "=SUM(D2:D114)"
    assert totals["Total"] == "=SUM(F2:F114)"
    assert worksheet.auto_filter.ref == "A1:F114"

    # Повторный запуск по тем же данным не увеличивает отчёт.
    assert main(
        [
            "--input",
            str(SAMPLE_SOURCE),
            "--template",
            str(SAMPLE_TEMPLATE),
            "--output",
            str(output),
            "--log",
            str(log_path),
        ]
    ) == 0
    rerun_book, rerun_sheet, rerun_rows, _rerun_totals = load_report(output)
    assert len(rerun_rows) == 113
    assert rgb(rerun_sheet.cell(1, 1).fill.fgColor).endswith("1F4E79")
    assert rerun_book["Справка"]["A1"].value == "Справка по шаблону"
    rerun_book.close()
    workbook.close()

    log_text = log_path.read_text(encoding="utf-8")
    assert "отрицательное количество" in log_text
    assert "некорректная дата" in log_text
    assert "дубликат" in log_text
    assert "Лист «Черновик»: отсутствуют обязательные колонки" in log_text
    assert "Traceback" not in log_text
