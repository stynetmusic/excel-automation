"""Проверка строк, расчёт Total и отбор дубликатов."""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any

from openpyxl.utils.datetime import from_excel

from config import MONEY_PRECISION, QUANTITY_PRECISION, TOTAL_TOLERANCE


@dataclass(frozen=True)
class RawRow:
    """Одна строка исходного листа. Значения ещё не проверены."""

    sheet: str
    row_number: int
    values: dict[str, Any]


@dataclass(frozen=True)
class ValidRecord:
    """Строка, которую можно переносить в итоговый отчёт."""

    date: date
    product: str
    sku: str
    quantity: Decimal
    unit_price: Decimal
    total: Decimal
    source_sheet: str
    source_row: int


@dataclass(frozen=True)
class Issue:
    """Сообщение для журнала: ошибка листа, некорректная строка или дубликат."""

    level: str
    sheet: str
    row_number: int | None
    message: str

    def format(self) -> str:
        if self.row_number is None:
            return f"Лист «{self.sheet}»: {self.message}"
        return f"Лист «{self.sheet}», строка {self.row_number}: {self.message}"


@dataclass
class ProcessResult:
    records: list[ValidRecord] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    rows_read: int = 0
    empty_rows: int = 0
    skipped: int = 0
    duplicates: int = 0

    @property
    def processed(self) -> int:
        return len(self.records)


_DATE_FORMATS = (
    "%Y-%m-%d",
    "%d.%m.%Y",
    "%d/%m/%Y",
    "%Y/%m/%d",
    "%d-%m-%Y",
    "%Y-%m-%d %H:%M:%S",
    "%d.%m.%Y %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
)


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def row_is_empty(values: dict[str, Any]) -> bool:
    return all(is_blank(value) for value in values.values())


def format_decimal(value: Decimal) -> str:
    """Короткий вид числа для текста журнала."""
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


# Пробелы, которыми в тексте отделяют тысячи: обычный, неразрывный,
# узкий неразрывный, цифровой и тонкий.
_THOUSAND_SPACES = (" ", "\u00a0", "\u202f", "\u2007", "\u2009")


def parse_decimal(value: Any) -> Decimal:
    """Преобразует число Excel или текст в Decimal.

    Текстовое правило, без попытки угадать неоднозначную запись:
    - одна точка или одна запятая — десятичный знак: ``1234.56``, ``1234,56``,
      ``0.99``, ``0,99``;
    - если есть и точка, и запятая, десятичный знак — тот, что стоит правее,
      другой отделяет тысячи группами по три цифры: ``1,234.56``, ``1.234,56``;
    - пробелы, в том числе неразрывные, отделяют только тысячи: ``1 234,56``;
    - несколько одинаковых точек или запятых без второго знака — тоже тысячи
      группами по три: ``1.234.567``. Запись ``1,23,45`` отклоняется;
    - голые ``1,234`` и ``1.234`` — это дробь 1.234, а не 1234. Тысячу в таком
      виде пишут как ``1 234``, ``1,234.00`` или ``1.234,00``;
    - минус сохраняется. Нечисло, в том числе ``abc``, даёт ошибку, а не ноль.
    """
    if isinstance(value, bool) or value is None:
        raise ValueError("не число")

    if isinstance(value, Decimal):
        number = value
    elif isinstance(value, int):
        number = Decimal(value)
    elif isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("не число")
        # Без промежуточного округления: точность задают Quantity и денежные поля.
        number = Decimal(str(value))
    elif isinstance(value, str):
        number = _parse_decimal_text(value)
    else:
        raise ValueError("не число")

    if not number.is_finite():
        raise ValueError("не число")
    return number


def _parse_decimal_text(value: str) -> Decimal:
    text = value.strip()
    for space in _THOUSAND_SPACES[1:]:
        text = text.replace(space, " ")
    text = text.strip()
    if text == "":
        raise ValueError("не число")

    sign = ""
    if text[0] in "+-":
        sign = text[0]
        text = text[1:].strip()
    if text == "" or any(character not in "0123456789., " for character in text):
        raise ValueError("не число")

    dots = text.count(".")
    commas = text.count(",")
    spaces = text.count(" ")

    if dots and commas:
        if spaces:
            raise ValueError("не число")
        decimal_sep = "." if text.rfind(".") > text.rfind(",") else ","
        thousands_sep = "," if decimal_sep == "." else "."
        if text.count(decimal_sep) != 1:
            raise ValueError("не число")
        integer, fraction = text.split(decimal_sep)
        digits = _digits_with_thousands(integer, thousands_sep)
        if not fraction.isdigit():
            raise ValueError("не число")
        normalized = f"{sign}{digits}.{fraction}"
    elif spaces and (dots or commas):
        decimal_sep = "." if dots else ","
        if text.count(decimal_sep) != 1:
            raise ValueError("не число")
        integer, fraction = text.split(decimal_sep)
        if not fraction.isdigit():
            raise ValueError("не число")
        digits = _digits_with_thousands(integer, " ")
        normalized = f"{sign}{digits}.{fraction}"
    elif spaces:
        normalized = f"{sign}{_digits_with_thousands(text, ' ')}"
    elif dots + commas == 1:
        decimal_sep = "." if dots else ","
        integer, fraction = text.split(decimal_sep)
        if not integer.isdigit() or not fraction.isdigit():
            raise ValueError("не число")
        normalized = f"{sign}{integer}.{fraction}"
    elif dots + commas > 1:
        thousands_sep = "." if dots else ","
        normalized = f"{sign}{_digits_with_thousands(text, thousands_sep)}"
    elif not text.isdigit():
        raise ValueError("не число")
    else:
        normalized = f"{sign}{text}"

    try:
        return Decimal(normalized)
    except InvalidOperation as exc:
        raise ValueError("не число") from exc


def _digits_with_thousands(integer: str, separator: str) -> str:
    """Проверяет группы тысяч и возвращает цифры без разделителя.

    Первая группа — от 1 до 3 цифр, каждая следующая — ровно 3.
    """
    if integer == "" or separator not in integer:
        raise ValueError("не число")
    parts = integer.split(separator)
    if any(part == "" or not part.isdigit() for part in parts):
        raise ValueError("не число")
    if not 1 <= len(parts[0]) <= 3:
        raise ValueError("не число")
    if any(len(part) != 3 for part in parts[1:]):
        raise ValueError("не число")
    return "".join(parts)


def parse_date(value: Any) -> date:
    """Принимает дату Excel, дату со временем и несколько текстовых форматов."""
    if isinstance(value, datetime):
        parsed = value.date()
    elif isinstance(value, date):
        parsed = value
    elif isinstance(value, bool):
        raise ValueError("не дата")
    elif isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("не дата")
        try:
            converted = from_excel(value)
        except (ValueError, OverflowError, TypeError) as exc:
            raise ValueError("не дата") from exc
        parsed = converted.date() if isinstance(converted, datetime) else converted
        if not isinstance(parsed, date):
            raise ValueError("не дата")
    elif isinstance(value, str):
        text = value.strip()
        parsed = None
        for fmt in _DATE_FORMATS:
            try:
                parsed = datetime.strptime(text, fmt).date()
                break
            except ValueError:
                continue
        if parsed is None:
            raise ValueError("не дата")
    else:
        raise ValueError("не дата")

    if parsed.year < 1900 or parsed.year > 2100:
        raise ValueError("не дата")
    return parsed


def _normalize_quantity(value: Decimal) -> Decimal:
    return value.quantize(QUANTITY_PRECISION, rounding=ROUND_HALF_UP)


def _normalize_money(value: Decimal) -> Decimal:
    return value.quantize(MONEY_PRECISION, rounding=ROUND_HALF_UP)


def _issue(row: RawRow, message: str, level: str = "warning") -> Issue:
    return Issue(level=level, sheet=row.sheet, row_number=row.row_number, message=message)


def validate_row(row: RawRow) -> tuple[ValidRecord | None, Issue | None]:
    """Проверяет одну непустую строку. Первая ошибка не прерывает остальные проверки полей."""
    reasons: list[str] = []
    values = row.values

    parsed_date: date | None = None
    date_value = values.get("Date")
    if is_blank(date_value):
        reasons.append("не заполнено поле Date")
    else:
        try:
            parsed_date = parse_date(date_value)
        except ValueError:
            reasons.append(f"некорректная дата: {date_value!r}")

    product: str | None = None
    product_value = values.get("Product")
    if is_blank(product_value):
        reasons.append("не заполнено поле Product")
    else:
        product = str(product_value).strip()

    sku: str | None = None
    sku_value = values.get("SKU")
    if is_blank(sku_value):
        reasons.append("не заполнено поле SKU")
    else:
        sku = str(sku_value).strip()

    quantity: Decimal | None = None
    quantity_value = values.get("Quantity")
    if is_blank(quantity_value):
        reasons.append("не заполнено поле Quantity")
    else:
        try:
            quantity = _normalize_quantity(parse_decimal(quantity_value))
        except ValueError:
            reasons.append(f"некорректное количество: {quantity_value!r}")
        else:
            if quantity < 0:
                reasons.append(f"отрицательное количество: {format_decimal(quantity)}")
                quantity = None

    price: Decimal | None = None
    price_value = values.get("Unit Price")
    if is_blank(price_value):
        reasons.append("не заполнено поле Unit Price")
    else:
        try:
            price = _normalize_money(parse_decimal(price_value))
        except ValueError:
            reasons.append(f"некорректная цена: {price_value!r}")
        else:
            if price < 0:
                reasons.append(f"отрицательная цена: {format_decimal(price)}")
                price = None

    total: Decimal | None = None
    total_value = values.get("Total")
    total_is_blank = is_blank(total_value)
    if not total_is_blank:
        try:
            total = parse_decimal(total_value)
        except ValueError:
            reasons.append(f"некорректное значение Total: {total_value!r}")

    if reasons or parsed_date is None or product is None or sku is None or quantity is None or price is None:
        message = "; ".join(reasons) if reasons else "строка не прошла проверку"
        return None, _issue(row, message)

    expected = _normalize_money(quantity * price)
    if total is None:
        stored_total = expected
    elif abs(total - expected) > TOTAL_TOLERANCE:
        return None, _issue(
            row,
            (
                f"Total {format_decimal(total)} не совпадает с расчётом "
                f"Quantity × Unit Price = {format(expected, '.2f')} "
                f"(допуск {format_decimal(TOTAL_TOLERANCE)})"
            ),
        )
    else:
        stored_total = _normalize_money(total)

    record = ValidRecord(
        date=parsed_date,
        product=product,
        sku=sku,
        quantity=quantity,
        unit_price=price,
        total=stored_total,
        source_sheet=row.sheet,
        source_row=row.row_number,
    )
    return record, None


def duplicate_key(record: ValidRecord) -> tuple[str, str, str]:
    """Уникальность: дата без времени, SKU и название товара после обрезки пробелов.

    Сравнение SKU и Product чувствительно к регистру.
    """
    return (record.date.isoformat(), record.sku, record.product)


def process_rows(rows: list[RawRow]) -> ProcessResult:
    """Объединяет уже прочитанные строки. Одна плохая строка не останавливает обработку.

    Дубликатом считается повтор уже принятой записи. Если первое вхождение
    некорректно, следующая корректная запись с тем же ключом сохраняется.
    """
    result = ProcessResult()
    seen: dict[tuple[str, str, str], ValidRecord] = {}

    for row in rows:
        if row_is_empty(row.values):
            result.empty_rows += 1
            continue

        result.rows_read += 1
        record, issue = validate_row(row)
        if issue is not None or record is None:
            result.skipped += 1
            if issue is not None:
                result.issues.append(issue)
            continue

        key = duplicate_key(record)
        first = seen.get(key)
        if first is not None:
            result.duplicates += 1
            result.issues.append(
                Issue(
                    level="warning",
                    sheet=row.sheet,
                    row_number=row.row_number,
                    message=(
                        "дубликат (Date={date}, SKU={sku}, Product={product}); "
                        "первая запись: лист «{sheet}», строка {row}"
                    ).format(
                        date=record.date.isoformat(),
                        sku=record.sku,
                        product=record.product,
                        sheet=first.source_sheet,
                        row=first.source_row,
                    ),
                )
            )
            continue

        seen[key] = record
        result.records.append(record)

    return result
