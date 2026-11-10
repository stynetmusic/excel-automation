"""Запуск объединения Excel-листов в итоговый отчёт.

Команда по умолчанию: python main.py
Исходный файл можно заменить: python main.py --input путь/к/файлу.xlsx
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from config import (
    DEFAULT_LOG,
    DEFAULT_OUTPUT,
    DEFAULT_SOURCE,
    DEFAULT_TEMPLATE,
    LOGGER_NAME,
    REQUIRED_COLUMNS,
)
from excel_handler import (
    ExcelAutomationError,
    NoDataSheetsError,
    read_source,
    write_report,
)
from logger import setup_logging
from processor import ProcessResult, process_rows


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Объединяет строки со всех листов исходной книги Excel "
            "и записывает проверенные данные в итоговый файл по шаблону."
        )
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=DEFAULT_SOURCE,
        help="Путь к исходной книге .xlsx",
    )
    parser.add_argument(
        "-t",
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help="Путь к шаблону итогового отчёта .xlsx",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help="Куда сохранить итоговый файл .xlsx",
    )
    parser.add_argument(
        "--log",
        type=Path,
        default=DEFAULT_LOG,
        help="Куда записывать журнал",
    )
    return parser.parse_args(argv)


def build_summary(
    result: ProcessResult,
    *,
    sheets_read: int,
    sheet_errors: int,
    empty_rows: int,
    output_path: Path,
    log_path: Path,
) -> str:
    rows = [
        ("Листов с данными", sheets_read),
        ("Прочитано строк", result.rows_read),
        ("Успешно обработано", result.processed),
        ("Пропущено записей", result.skipped),
        ("Обнаружено дубликатов", result.duplicates),
        ("Пустых строк пропущено", empty_rows),
        ("Листов с ошибкой структуры", sheet_errors),
        ("Итоговый файл", output_path),
        ("Журнал ошибок", log_path),
    ]
    width = max(len(label) for label, _value in rows)
    lines = ["Обработка завершена."]
    lines.extend(f"{label:<{width}}  {value}" for label, value in rows)
    if result.skipped or result.duplicates or sheet_errors:
        lines.append("Некорректные строки и дубликаты перечислены в журнале.")
    return "\n".join(lines)


def run(source: Path, template: Path, output: Path, log_path: Path) -> int:
    logger = logging.getLogger(LOGGER_NAME)
    source = source.expanduser().resolve()
    template = template.expanduser().resolve()
    output = output.expanduser().resolve()
    log_path = log_path.expanduser().resolve()

    logger.info("Запуск обработки")
    logger.info("Источник: %s", source)
    logger.info("Шаблон: %s", template)
    logger.info("Итог: %s", output)

    if source == output:
        raise ExcelAutomationError("Исходный файл и итоговый файл не должны совпадать.")
    if template == output:
        raise ExcelAutomationError("Шаблон и итоговый файл не должны совпадать.")

    read_result = read_source(source)
    for issue in read_result.issues:
        logger.log(_level_number(issue.level), "%s", issue.format())

    if not read_result.sheets_read:
        columns = ", ".join(REQUIRED_COLUMNS)
        raise NoDataSheetsError(
            "В книге не найдено ни одного листа с колонками: " + columns
        )

    for sheet in read_result.sheets_read:
        count = sum(1 for row in read_result.rows if row.sheet == sheet)
        logger.info("Прочитан лист «%s»: строк на листе, включая пустые: %s", sheet, count)

    processed = process_rows(read_result.rows)
    for issue in processed.issues:
        logger.log(_level_number(issue.level), "%s", issue.format())

    write_report(processed.records, template, output)

    sheet_errors = sum(1 for issue in read_result.issues if issue.level == "error")
    summary = build_summary(
        processed,
        sheets_read=len(read_result.sheets_read),
        sheet_errors=sheet_errors,
        empty_rows=processed.empty_rows,
        output_path=output,
        log_path=log_path,
    )
    logger.info(
        "Итог: прочитано %s, обработано %s, пропущено %s, дубликатов %s, пустых %s",
        processed.rows_read,
        processed.processed,
        processed.skipped,
        processed.duplicates,
        processed.empty_rows,
    )
    print()
    print(summary)
    return 0


def _level_number(level: str) -> int:
    return int(getattr(logging, level.upper(), logging.WARNING))


def main(argv: list[str] | None = None) -> int:
    if sys.version_info < (3, 11):
        print("Нужен Python 3.11 или новее.")
        return 1

    args = parse_args(argv)
    log_path = args.log.expanduser().resolve()
    try:
        setup_logging(log_path)
    except OSError as exc:
        print(f"Ошибка: не удалось открыть журнал {log_path}: {exc}")
        return 1

    logger = logging.getLogger(LOGGER_NAME)
    try:
        return run(args.input, args.template, args.output, log_path)
    except ExcelAutomationError as exc:
        logger.error("%s", exc)
        print(f"Ошибка: {exc}")
        print(f"Журнал ошибок: {log_path}")
        return 1
    except Exception:
        logger.exception("Непредвиденная ошибка")
        print("Непредвиденная ошибка. Подробности записаны в журнал.")
        print(f"Журнал ошибок: {log_path}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
