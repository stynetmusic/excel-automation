"""Проверка правил строк без Excel: расчёт, допуск, пустые значения и дубликаты."""

from datetime import date

import pytest

from config import REQUIRED_COLUMNS
from processor import RawRow, duplicate_key, parse_date, parse_decimal, process_rows


def raw(sheet: str, row_number: int, **values: object) -> RawRow:
    payload = {name: None for name in REQUIRED_COLUMNS}
    payload.update(values)
    return RawRow(sheet=sheet, row_number=row_number, values=payload)


def good(**overrides: object) -> dict[str, object]:
    values: dict[str, object] = {
        "Date": date(2026, 4, 1),
        "Product": "Чай",
        "SKU": "TEA-1",
        "Quantity": 2,
        "Unit Price": "10.00",
        "Total": "20.00",
    }
    values.update(overrides)
    return values


def test_calculates_missing_total() -> None:
    result = process_rows([raw("North", 2, **good(Total=None, Quantity=4, **{"Unit Price": "9.99"}))])
    assert result.processed == 1
    assert result.skipped == 0
    assert result.records[0].total == parse_decimal("39.96")


def test_accepts_exact_total_and_text_numbers() -> None:
    result = process_rows(
        [raw("North", 2, **good(Quantity="3", **{"Unit Price": "12,50"}, Total="37.50"))]
    )
    assert result.processed == 1
    assert result.records[0].quantity == parse_decimal("3")
    assert result.records[0].unit_price == parse_decimal("12.50")
    assert result.records[0].total == parse_decimal("37.50")


def test_tolerance_boundaries() -> None:
    accepted = process_rows(
        [raw("North", 2, **good(Quantity=1, **{"Unit Price": "35.50"}, Total="35.51"))]
    )
    rejected = process_rows(
        [raw("North", 2, **good(Quantity=1, **{"Unit Price": "10.00"}, Total="10.02"))]
    )
    rounded = process_rows(
        [raw("North", 2, **good(Quantity=3, **{"Unit Price": "1.00"}, Total="3.009"))]
    )
    assert accepted.processed == 1
    assert accepted.records[0].total == parse_decimal("35.51")
    assert rejected.processed == 0
    assert rejected.skipped == 1
    assert "не совпадает" in rejected.issues[0].message
    assert rounded.processed == 1
    assert rounded.records[0].total == parse_decimal("3.01")


def test_blank_total_text_is_calculated() -> None:
    result = process_rows([raw("North", 2, **good(Total="  ", Quantity=2, **{"Unit Price": "8.40"}))])
    assert result.processed == 1
    assert result.records[0].total == parse_decimal("16.80")


def test_skips_empty_and_whitespace_rows() -> None:
    rows = [
        raw("North", 2, **good(SKU="A")),
        raw("North", 3),
        raw("North", 4, Date=" ", Product=" ", SKU=" ", Quantity=" ", **{"Unit Price": " ", "Total": " "}),
        raw("North", 5, **good(SKU="B")),
    ]
    result = process_rows(rows)
    assert result.empty_rows == 2
    assert result.rows_read == 2
    assert [record.sku for record in result.records] == ["A", "B"]
    assert result.issues == []


def test_invalid_values_do_not_stop_the_batch() -> None:
    rows = [
        raw("North", 2, **good(Date="вчера", SKU="BAD-DATE")),
        raw("North", 3, **good(Quantity=-3, SKU="BAD-QTY")),
        raw("North", 4, **good(**{"Unit Price": -1}, SKU="BAD-PRICE")),
        raw("North", 5, **good(Quantity="много", SKU="TEXT-QTY")),
        raw("North", 6, **good(**{"Unit Price": "N/A"}, SKU="TEXT-PRICE")),
        raw("North", 7, **good(Product="  ", SKU="NO-NAME")),
        raw("North", 8, **good(SKU="GOOD", Quantity=0, Total=0, **{"Unit Price": 0})),
    ]
    result = process_rows(rows)
    assert result.processed == 1
    assert result.records[0].sku == "GOOD"
    assert result.skipped == 6
    messages = " | ".join(issue.message for issue in result.issues)
    assert "некорректная дата" in messages
    assert "отрицательное количество" in messages
    assert "отрицательная цена" in messages
    assert "некорректное количество" in messages
    assert "некорректная цена" in messages
    assert "не заполнено поле Product" in messages


def test_several_problems_in_one_row_count_as_one_skip() -> None:
    result = process_rows(
        [raw("North", 9, **good(Date="32.13.2026", Product=None, Quantity=-1, SKU="MULTI"))]
    )
    assert result.skipped == 1
    assert result.processed == 0
    assert "некорректная дата" in result.issues[0].message
    assert "не заполнено поле Product" in result.issues[0].message
    assert "отрицательное количество" in result.issues[0].message


def test_duplicates_keep_the_first_valid_record() -> None:
    rows = [
        raw("North", 2, **good(SKU="DUP", Quantity=3, Total="30.00")),
        raw("South", 4, **good(SKU="DUP", Quantity=9, Total="90.00")),
        raw("North", 5, **good(SKU="DUP", Product="Другой товар", Quantity=1, Total="10.00")),
        raw("East", 2, **good(Date="не дата", SKU="LATER", Quantity=1)),
        raw("East", 3, **good(SKU="LATER", Quantity=4, Total="40.00")),
    ]
    result = process_rows(rows)
    kept = {(record.sku, record.product): record for record in result.records}
    assert kept[("DUP", "Чай")].quantity == parse_decimal("3")
    assert kept[("DUP", "Чай")].source_sheet == "North"
    assert ("DUP", "Другой товар") in kept
    assert kept[("LATER", "Чай")].quantity == parse_decimal("4")
    assert result.duplicates == 1
    assert "первая запись: лист «North», строка 2" in result.issues[-1].message or any(
        "строка 2" in issue.message for issue in result.issues if "дубликат" in issue.message
    )


def test_duplicate_key_is_case_sensitive_and_ignores_time() -> None:
    from datetime import datetime

    first = process_rows([raw("North", 2, **good())]).records[0]
    second = process_rows(
        [raw("North", 3, **good(Date=datetime(2026, 4, 1, 18, 30), Product="чай"))]
    ).records[0]
    assert duplicate_key(first)[0] == "2026-04-01"
    assert duplicate_key(first) != duplicate_key(second)


def test_date_formats() -> None:
    assert parse_date("2026-03-15") == date(2026, 3, 15)
    assert parse_date("15.03.2026") == date(2026, 3, 15)
    assert parse_date("15/03/2026") == date(2026, 3, 15)
    assert parse_date("2026-03-15 09:45:00") == date(2026, 3, 15)
    with pytest.raises(ValueError):
        parse_date("03/15/2026")
    with pytest.raises(ValueError):
        parse_date("вчера")


def test_strips_product_and_sku() -> None:
    result = process_rows(
        [raw("Online", 2, **good(Product="  Керамика  ", SKU="  MUG-1  ", Total="20.00"))]
    )
    assert result.records[0].product == "Керамика"
    assert result.records[0].sku == "MUG-1"
