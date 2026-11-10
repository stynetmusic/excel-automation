"""Пути, колонки и постоянные правила обработки."""

from decimal import Decimal
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
INPUT_DIR = ROOT_DIR / "input"
TEMPLATE_DIR = ROOT_DIR / "templates"
OUTPUT_DIR = ROOT_DIR / "output"
LOG_DIR = ROOT_DIR / "logs"

DEFAULT_SOURCE = INPUT_DIR / "source_sales.xlsx"
DEFAULT_TEMPLATE = TEMPLATE_DIR / "report_template.xlsx"
DEFAULT_OUTPUT = OUTPUT_DIR / "consolidated_report.xlsx"
DEFAULT_LOG = LOG_DIR / "app.log"

# Заголовки исходных листов и итогового шаблона. Регистр и лишние пробелы
# при чтении не учитываются, в файл записываются эти названия.
REQUIRED_COLUMNS = (
    "Date",
    "Product",
    "SKU",
    "Quantity",
    "Unit Price",
    "Total",
)

REPORT_SHEET_NAME = "Сводный отчёт"
HELP_SHEET_NAME = "Справка"

# Допуск сверки Total с произведением Quantity × Unit Price.
TOTAL_TOLERANCE = Decimal("0.01")

# Количество хранится до 4 знаков, деньги — до 2 знаков.
QUANTITY_PRECISION = Decimal("0.0001")
MONEY_PRECISION = Decimal("0.01")

HEADER_ROW = 1
LOGGER_NAME = "excel_automation"
