"""Запуск программы: ошибки файлов, журнал и команда из другого каталога."""

from __future__ import annotations

import os
import subprocess
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from config import REQUIRED_COLUMNS
from main import main
from tests.helpers import load_report, make_source, make_template

ROOT = Path(__file__).resolve().parents[1]


def _args(source: Path, template: Path, output: Path, log_path: Path) -> list[str]:
    return [
        "--input",
        str(source),
        "--template",
        str(template),
        "--output",
        str(output),
        "--log",
        str(log_path),
    ]


def _summary_value(text: str, label: str) -> str:
    for line in text.splitlines():
        if line.startswith(label):
            return line[len(label):].strip()
    raise AssertionError(f"В сводке нет строки {label!r}:\n{text}")


def test_summary_counts_for_sample(capsys, tmp_path: Path) -> None:
    output = tmp_path / "out.xlsx"
    log_path = tmp_path / "app.log"
    code = main(
        _args(
            ROOT / "input" / "source_sales.xlsx",
            ROOT / "templates" / "report_template.xlsx",
            output,
            log_path,
        )
    )
    assert code == 0
    text = capsys.readouterr().out
    assert _summary_value(text, "Листов с данными") == "6"
    assert _summary_value(text, "Прочитано строк") == "128"
    assert _summary_value(text, "Успешно обработано") == "113"
    assert _summary_value(text, "Пропущено записей") == "12"
    assert _summary_value(text, "Обнаружено дубликатов") == "3"
    assert _summary_value(text, "Пустых строк пропущено") == "3"
    assert _summary_value(text, "Листов с ошибкой структуры") == "1"
    assert _summary_value(text, "Итоговый файл") == str(output.resolve())
    assert _summary_value(text, "Журнал ошибок") == str(log_path.resolve())
    assert "Некорректные строки и дубликаты перечислены в журнале." in text


def test_missing_corrupt_and_bad_paths(capsys, tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    log_path = tmp_path / "app.log"
    make_template(template)

    missing = main(_args(tmp_path / "нет-такого.xlsx", template, output, log_path))
    missing_text = capsys.readouterr().out
    assert missing == 1
    assert "не найден" in missing_text
    assert "Traceback" not in missing_text
    assert "не найден" in log_path.read_text(encoding="utf-8")

    corrupt = tmp_path / "broken.xlsx"
    corrupt.write_text("это не книга Excel", encoding="utf-8")
    corrupt_code = main(_args(corrupt, template, output, log_path))
    corrupt_text = capsys.readouterr().out
    assert corrupt_code == 1
    assert "повреждён" in corrupt_text
    assert "Traceback" not in corrupt_text

    wrong_type = tmp_path / "book.xls"
    wrong_type.write_bytes(b"not-xlsx")
    wrong_code = main(_args(wrong_type, template, output, log_path))
    assert wrong_code == 1
    assert ".xlsx" in capsys.readouterr().out

    empty_book = tmp_path / "empty.xlsx"
    workbook = Workbook()
    workbook.active.title = "Пусто"
    workbook.save(empty_book)
    workbook.close()
    empty_code = main(_args(empty_book, template, output, log_path))
    assert empty_code == 1
    assert "не найдено" in capsys.readouterr().out

    source = tmp_path / "one.xlsx"
    make_source(
        source,
        [("Лист", list(REQUIRED_COLUMNS), [{"Date": date(2026, 6, 1), "Product": "А", "SKU": "A", "Quantity": 1, "Unit Price": 5, "Total": 5}])],
    )
    same = main(_args(source, template, source, log_path))
    assert same == 1
    assert "не должны совпадать" in capsys.readouterr().out

    missing_template = main(_args(source, tmp_path / "no-template.xlsx", output, log_path))
    assert missing_template == 1
    assert "Шаблон не найден" in capsys.readouterr().out


def test_unwritable_output_is_reported(tmp_path: Path, capsys) -> None:
    source = tmp_path / "source.xlsx"
    template = tmp_path / "template.xlsx"
    locked = tmp_path / "locked"
    locked.mkdir()
    make_source(
        source,
        [("Лист", list(REQUIRED_COLUMNS), [{"Date": date(2026, 6, 1), "Product": "А", "SKU": "A", "Quantity": 1, "Unit Price": 5, "Total": 5}])],
    )
    make_template(template)
    os.chmod(locked, 0o555)
    try:
        code = main(_args(source, template, locked / "out.xlsx", tmp_path / "app.log"))
        text = capsys.readouterr().out
    finally:
        os.chmod(locked, 0o755)
    assert code == 1
    assert "Не удалось сохранить" in text or "Нет доступа" in text
    assert "Traceback" not in text


def test_one_bad_row_does_not_abort(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    template = tmp_path / "template.xlsx"
    output = tmp_path / "out.xlsx"
    make_source(
        source,
        [
            (
                "Лист",
                list(REQUIRED_COLUMNS),
                [
                    {"Date": "вчера", "Product": "Плохая", "SKU": "BAD", "Quantity": 1, "Unit Price": 5, "Total": 5},
                    {"Date": date(2026, 6, 2), "Product": "Хорошая", "SKU": "GOOD", "Quantity": 2, "Unit Price": 3, "Total": None},
                ],
            )
        ],
    )
    make_template(template)
    assert main(_args(source, template, output, tmp_path / "app.log")) == 0
    workbook, _worksheet, rows, _totals = load_report(output)
    assert [row["SKU"] for row in rows] == ["GOOD"]
    assert abs(float(rows[0]["Total"]) - 6) < 0.001
    workbook.close()


def test_cli_from_another_directory(tmp_path: Path) -> None:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    completed = subprocess.run(
        [
            sys.executable,
            str(ROOT / "main.py"),
            "--input",
            str(ROOT / "input" / "source_sales.xlsx"),
            "--template",
            str(ROOT / "templates" / "report_template.xlsx"),
            "--output",
            str(tmp_path / "from-other-dir.xlsx"),
            "--log",
            str(tmp_path / "app.log"),
        ],
        cwd="/tmp",
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    assert "Traceback" not in completed.stderr
    assert "Успешно обработано" in completed.stdout
    assert "113" in completed.stdout
    workbook, _worksheet, rows, _totals = load_report(tmp_path / "from-other-dir.xlsx")
    assert len(rows) == 113
    workbook.close()


def test_default_paths_when_launched_outside_project(tmp_path: Path) -> None:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    completed = subprocess.run(
        [sys.executable, str(ROOT / "main.py")],
        cwd=str(tmp_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    output = ROOT / "output" / "consolidated_report.xlsx"
    log_path = ROOT / "logs" / "app.log"
    assert output.exists()
    assert log_path.exists()
    assert str(output) in completed.stdout
    workbook, worksheet, rows, totals = load_report(output)
    assert len(rows) == 113
    assert totals is not None
    assert worksheet.cell(1, 1).value == "Date"
    assert workbook["Справка"]["A1"].value == "Справка по шаблону"
    workbook.close()
    second = subprocess.run(
        [sys.executable, str(ROOT / "main.py")],
        cwd=str(tmp_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
        check=False,
    )
    assert second.returncode == 0, second.stderr
    workbook, _worksheet, rows_again, _totals = load_report(output)
    assert len(rows_again) == 113
    workbook.close()
