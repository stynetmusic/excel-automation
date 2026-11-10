"""Текстовые числа: европейская и американская запись, знак и отказ от нуля вместо мусора."""

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from config import REQUIRED_COLUMNS
from main import main
from processor import RawRow, parse_decimal, process_rows
from tests.helpers import assert_number, load_report, make_source, make_template

NBSP = "\u00a0"


def raw(**values: object) -> RawRow:
    payload = {name: None for name in REQUIRED_COLUMNS}
    payload.update(
        {
            "Date": date(2026, 7, 1),
            "Product": "Число",
            "SKU": "NUM",
            "Quantity": 1,
            "Unit Price": "1.00",
            "Total": None,
        }
    )
    payload.update(values)
    return RawRow(sheet="Лист", row_number=2, values=payload)


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("1234.56", "1234.56"),
        ("1234,56", "1234.56"),
        ("1.234,56", "1234.56"),
        ("1,234.56", "1234.56"),
        ("1 234,56", "1234.56"),
        (f"1{NBSP}234,56", "1234.56"),
        ("0,99", "0.99"),
        ("0.99", "0.99"),
        ("-123,45", "-123.45"),
    ],
)
def test_localized_numbers_parse_exactly(text: str, expected: str) -> None:
    assert parse_decimal(text) == Decimal(expected)


@pytest.mark.parametrize("text", ["abc", "1,23,45"])
def test_invalid_numbers_are_rejected_and_not_replaced_with_zero(text: str) -> None:
    with pytest.raises(ValueError):
        parse_decimal(text)

    result = process_rows([raw(SKU=text, Quantity=text, **{"Unit Price": "10.00"})])
    assert result.processed == 0
    assert result.skipped == 1
    assert result.records == []
    assert text in result.issues[0].message


def test_quantity_float_rounds_once_to_four_places() -> None:
    result = process_rows([raw(Quantity=1.2345499, SKU="Q-BORDER")])
    assert result.processed == 1
    assert result.records[0].quantity == Decimal("1.2345")


def test_unit_price_float_rounds_once_to_two_places() -> None:
    result = process_rows([raw(**{"Unit Price": 1.2349999}, SKU="P-BORDER")])
    assert result.processed == 1
    assert result.records[0].unit_price == Decimal("1.23")
    assert result.records[0].total == Decimal("1.23")


@pytest.mark.parametrize("space", ["\u202f", "\u2007", "\u2009"])
def test_narrow_spaces_are_thousand_separators(space: str) -> None:
    assert parse_decimal(f"1{space}234,56") == Decimal("1234.56")


@pytest.mark.parametrize(
    "value",
    [float("nan"), float("inf"), float("-inf"), True, False, "abc", "1,23,45"],
)
def test_non_finite_bool_and_bad_text_are_rejected(value: object) -> None:
    with pytest.raises(ValueError):
        parse_decimal(value)

    result = process_rows([raw(Quantity=value, SKU="REJECT")])
    assert result.processed == 0
    assert result.records == []
    assert result.skipped == 1


def test_negative_sign_is_preserved_for_later_validation() -> None:
    assert parse_decimal("-123,45") == Decimal("-123.45")
    assert parse_decimal("-123,45") < 0

    result = process_rows([raw(Quantity="-123,45", **{"Unit Price": "1.00"})])
    assert result.processed == 0
    assert result.skipped == 1
    assert "отрицательное количество: -123.45" in result.issues[0].message


def test_money_uses_decimal_not_binary_float() -> None:
    """1.234,56 × 0,99 = 1222.2144, до копеек 1222.21. Float здесь дал бы шум."""
    result = process_rows(
        [raw(Quantity="1.234,56", **{"Unit Price": "0,99"}, Total=None, SKU="MONEY")]
    )
    assert result.processed == 1
    record = result.records[0]
    assert record.quantity == Decimal("1234.5600")
    assert record.unit_price == Decimal("0.99")
    assert record.total == Decimal("1222.21")
    assert record.total == (record.quantity * record.unit_price).quantize(Decimal("0.01"))


def test_single_separator_is_decimal_by_documented_rule() -> None:
    assert parse_decimal("1,234") == Decimal("1.234")
    assert parse_decimal("1.234") == Decimal("1.234")
    assert parse_decimal("1 234") == Decimal("1234")
    assert parse_decimal("1.234.567,89") == Decimal("1234567.89")
    with pytest.raises(ValueError):
        parse_decimal("1,23,567.89")


def test_localized_text_in_excel_cells(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    log_path = tmp_path / "app.log"
    make_source(
        source,
        [
            (
                "Лист",
                list(REQUIRED_COLUMNS),
                [
                    {
                        "Date": date(2026, 7, 1),
                        "Product": "Европейская цена",
                        "SKU": "EU-1",
                        "Quantity": "1.234,56",
                        "Unit Price": "0,99",
                        "Total": None,
                    },
                    {
                        "Date": date(2026, 7, 2),
                        "Product": "Американская цена",
                        "SKU": "US-1",
                        "Quantity": 2,
                        "Unit Price": "1,234.56",
                        "Total": "2.469,12",
                    },
                    {
                        "Date": date(2026, 7, 3),
                        "Product": "Неразрывный пробел",
                        "SKU": "SP-1",
                        "Quantity": f"1{NBSP}234,50",
                        "Unit Price": "1.00",
                        "Total": None,
                    },
                    {
                        "Date": date(2026, 7, 4),
                        "Product": "Мусор",
                        "SKU": "BAD-ABC",
                        "Quantity": "abc",
                        "Unit Price": 1,
                        "Total": 1,
                    },
                    {
                        "Date": date(2026, 7, 5),
                        "Product": "Плохие группы",
                        "SKU": "BAD-GROUPS",
                        "Quantity": "1,23,45",
                        "Unit Price": "1.00",
                        "Total": None,
                    },
                    {
                        "Date": date(2026, 7, 6),
                        "Product": "Отрицательное",
                        "SKU": "NEG-TXT",
                        "Quantity": "-123,45",
                        "Unit Price": "1,00",
                        "Total": None,
                    },
                    {
                        "Date": date(2026, 7, 7),
                        "Product": "Число Excel",
                        "SKU": "NUM-1",
                        "Quantity": 2,
                        "Unit Price": 10.5,
                        "Total": 21,
                    },
                ],
            )
        ],
    )
    make_template(template)
    assert main(
        ["--input", str(source), "--template", str(template), "--output", str(output), "--log", str(log_path)]
    ) == 0

    workbook, _worksheet, rows, totals = load_report(output)
    by_sku = {row["SKU"]: row for row in rows}
    assert set(by_sku) == {"EU-1", "US-1", "SP-1", "NUM-1"}
    assert_number(by_sku["EU-1"]["Quantity"], 1234.56)
    assert_number(by_sku["EU-1"]["Unit Price"], 0.99)
    assert_number(by_sku["EU-1"]["Total"], 1222.21)
    assert_number(by_sku["US-1"]["Unit Price"], 1234.56)
    assert_number(by_sku["US-1"]["Total"], 2469.12)
    assert_number(by_sku["SP-1"]["Quantity"], 1234.50)
    assert_number(by_sku["SP-1"]["Total"], 1234.50)
    assert_number(by_sku["NUM-1"]["Total"], 21)
    assert totals is not None
    assert str(totals["Total"]).startswith("=SUM(")

    log_text = log_path.read_text(encoding="utf-8")
    assert "некорректное количество: 'abc'" in log_text
    assert "некорректное количество: '1,23,45'" in log_text
    assert "отрицательное количество: -123.45" in log_text
    assert "BAD-ABC" not in {row["SKU"] for row in rows}
    workbook.close()
