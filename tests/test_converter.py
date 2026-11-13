import csv
import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from mocod_converter.converter import HEADERS, ConversionError, convert


class ConverterTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.source = Path(directory.name) / "Реестр.xlsx"
        self.target = Path(directory.name) / "Реестр.csv"
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        for number in range(1, 121):
            self.sheet.cell(3, number, number)
        self.sheet["A7"] = 1

    def test_csv_layout_and_values(self):
        self.sheet["B7"] = 'Пример; "текст"'
        self.sheet["C7"] = "0001_000002"
        self.sheet["D7"] = date(2026, 1, 3)
        self.sheet["H7"] = 123
        self.sheet["H7"].number_format = "0000000000"
        self.sheet["P7"] = "Да"
        self.sheet["AB7"] = 123.45
        self.sheet["AJ7"] = 0.75
        self.sheet["AJ7"].number_format = "0.000%"
        self.sheet["BX7"] = "не рассчитывался"
        self.workbook.save(self.source)
        original = self.source.read_bytes()

        self.assertEqual(convert(self.source, self.target), 1)
        content = self.target.read_bytes()
        self.assertTrue(content.endswith(b"\r\n"))
        self.assertIn('"Пример; ""текст"""', content.decode("cp1251"))
        with self.target.open(encoding="cp1251", newline="") as stream:
            headers, row = csv.reader(stream, delimiter=";")
        self.assertEqual(tuple(headers), HEADERS)
        self.assertEqual(len(row), 120)
        for field, expected in {
            2: 'Пример; "текст"', 3: "0001_000002", 4: "2026-01-03",
            8: "0000000123", 16: "да", 28: "123,45", 36: "75,000%", 76: "",
        }.items():
            self.assertEqual(row[field - 1], expected)
        self.assertEqual(self.source.read_bytes(), original)

    def test_bad_encoding_preserves_existing_output(self):
        self.sheet["B7"] = "🙂"
        self.workbook.save(self.source)
        self.target.write_bytes(b"existing")
        with self.assertRaisesRegex(ConversionError, "F2"):
            convert(self.source, self.target)
        self.assertEqual(self.target.read_bytes(), b"existing")

    def test_missing_formula_result_identifies_field(self):
        self.sheet["AB7"] = "=1+2"
        self.workbook.save(self.source)
        with self.assertRaisesRegex(ConversionError, "F28"):
            convert(self.source, self.target)
        self.assertFalse(self.target.exists())

    def test_cli_supports_cyrillic_paths_and_protects_existing_output(self):
        self.workbook.save(self.source)
        command = [sys.executable, "-m", "mocod_converter", str(self.source), str(self.target)]
        first = subprocess.run(command, capture_output=True, check=False)
        self.assertEqual(first.returncode, 0, first.stderr)
        original = self.target.read_bytes()
        again = subprocess.run(command, capture_output=True, check=False)
        self.assertEqual(again.returncode, 1, again.stderr)
        self.assertEqual(self.target.read_bytes(), original)
