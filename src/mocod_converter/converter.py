import os
import re
from contextlib import ExitStack
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


HEADERS = tuple(f"F{i}" for i in range(1, 121))
NUMBERS = tuple(str(i) for i in range(1, 121))
DATE_FIELDS = {
    4, 5, 18, 23, 24, 31, 33, 34, 35, 42, 49, 50, 57, 58, 59, 60,
    63, 64, 67, 75, 90, 91, 92, 105, 116, 119, 120,
}
INTEGER_FIELDS = {1, 9, 32, 102, 103, 104}
AMOUNT_FIELDS = {
    19, 20, 21, 28, 29, 30, 40, 48, 68, 69, 70, 72, 74, 83, 84, 85,
    86, 87, 88, 89, *range(93, 102), *range(106, 116), 117, 118,
}
PERCENT_FIELDS = dict.fromkeys((36, 39, 41), 3) | dict.fromkeys(range(76, 81), 1)
TEXT_FIELDS = set(range(1, 121)) - DATE_FIELDS - INTEGER_FIELDS - AMOUNT_FIELDS - PERCENT_FIELDS.keys()


class ConversionError(ValueError):
    pass


def find_register(workbook):
    matches = []
    for sheet in workbook.worksheets:
        for row_number, row in enumerate(sheet.iter_rows(max_row=3, values_only=True), 1):
            labels = tuple("" if value is None else str(value).strip() for value in row)
            for first_column, label in enumerate(labels):
                if label not in ("1", "F1"):
                    continue
                if labels[first_column:first_column + 120] in (NUMBERS, HEADERS):
                    matches.append((sheet.title, first_column + 1, 7 if row_number == 3 else row_number + 1))
    if len(matches) != 1:
        if matches:
            raise ConversionError("В книге несколько таблиц реестра. Оставьте в копии книги один реестр.")
        raise ConversionError("Не найдены 120 последовательных граф реестра в строке заголовков.")
    return matches[0]


def as_date(value, epoch):
    if isinstance(value, (int, float)):
        value = from_excel(value, epoch)
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str):
        for pattern in ("%Y-%m-%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(value.strip(), pattern).date().isoformat()
            except ValueError:
                continue
    raise ValueError("ожидается дата в формате ГГГГ-ММ-ДД")


def as_decimal(value):
    cleaned = re.sub(r"[\u0020\u00a0\u202f]", "", str(value).strip())
    if not re.fullmatch(r"[+-]?\d+(?:[,.]\d+)?(?:[Ee][+-]?\d+)?", cleaned):
        raise ValueError("ожидается число")
    return Decimal(cleaned.replace(",", "."))


def field_value(cell, formula_cell, column, epoch):
    value = cell.value
    if value is None or value == "":
        if value is None and formula_cell.data_type == "f":
            raise ValueError("у формулы нет сохранённого результата. Пересчитайте и сохраните книгу в Excel")
        return ""
    if cell.data_type == "e":
        raise ValueError("ошибка в ячейке Excel")
    if column in DATE_FIELDS:
        return as_date(value, epoch)
    if column in PERCENT_FIELDS:
        if column >= 76 and isinstance(value, str) and value.strip().casefold() == "не рассчитывался":
            return ""
        places = PERCENT_FIELDS[column]
        number = as_decimal(str(value).strip().removesuffix("%"))
        if isinstance(value, (int, float)) and "%" in cell.number_format:
            number *= 100
        number = number.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
        return f"{number:.{places}f}".replace(".", ",") + "%"
    if column in INTEGER_FIELDS or column in AMOUNT_FIELDS:
        number = as_decimal(value)
        if column in INTEGER_FIELDS and number != number.to_integral_value():
            raise ValueError("ожидается целое число")
        if number == number.to_integral_value():
            return format(number.to_integral_value(), "f")
        return format(number.normalize(), "f").replace(".", ",")
    if isinstance(value, str):
        if column == 16 and value.casefold() == "да":
            return "да"
        return value
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = Decimal(str(value))
        if re.fullmatch("0+", cell.number_format) and number == number.to_integral_value():
            return format(number.to_integral_value(), "f").zfill(len(cell.number_format))
        return format(number.normalize(), "f")
    return str(value)


def convert(source, target):
    if source.suffix.lower() != ".xlsx" or target.suffix.lower() != ".csv":
        raise ConversionError("Выберите книгу XLSX и путь для файла CSV.")
    temporary = None
    try:
        with ExitStack() as stack:
            book = load_workbook(source, read_only=True, data_only=True, keep_links=False)
            stack.callback(book.close)
            formula_book = load_workbook(source, read_only=True, keep_links=False)
            stack.callback(formula_book.close)
            sheet, first_column, first_row = find_register(book)
            bounds = dict(min_row=first_row, min_col=first_column, max_col=first_column + 119)
            rows = book[sheet].iter_rows(**bounds)
            formula_rows = formula_book[sheet].iter_rows(**bounds)
            count = 0
            with NamedTemporaryFile(mode="w", encoding="cp1251", newline="", suffix=".tmp",
                                    dir=target.parent, delete=False) as stream:
                temporary = Path(stream.name)
                stream.write(";".join(HEADERS) + "\r\n")
                for row_number, (row, formula_row) in enumerate(zip(rows, formula_rows, strict=True), first_row):
                    values = []
                    for column, (cell, formula_cell) in enumerate(zip(row, formula_row, strict=True), 1):
                        try:
                            value = field_value(cell, formula_cell, column, book.epoch)
                            value.encode("cp1251")
                        except UnicodeEncodeError:
                            raise ConversionError(
                                f"Строка {row_number}, F{column}. Символ не поддерживается Windows-1251."
                            ) from None
                        except (ValueError, InvalidOperation, OverflowError) as error:
                            raise ConversionError(f"Строка {row_number}, F{column}. {error}.") from None
                        if value and column in TEXT_FIELDS:
                            value = '"' + value.replace('"', '""') + '"'
                        values.append(value)
                    if any(values) or any(cell.value not in (None, "") for cell in row):
                        stream.write(";".join(values) + "\r\n")
                        count += 1
            if not count:
                raise ConversionError("В реестре нет строк с данными.")
        os.replace(temporary, target)
        return count
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
