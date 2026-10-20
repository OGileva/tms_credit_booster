from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from qa_cases.core import (
    DEFAULT_HEADERS,
    DEFAULT_SHEET,
    RepositoryConfig,
    export_cases,
    load_config,
    load_suite,
    select_cases,
    validate_repository,
)


class RepositoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        (self.root / "cases" / "credit-booster" / "payments").mkdir(parents=True)
        (self.root / "suites").mkdir()
        (self.root / "project.json").write_text(
            json.dumps(
                {
                    "sheet_name": DEFAULT_SHEET,
                    "headers": DEFAULT_HEADERS,
                    "id_pattern": r"^(?P<id>[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+-\d{4})\s+—\s+\S",
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

    def tearDown(self):
        self.temp.cleanup()

    def _write_case(self, filename: str, case_id: str = "CB-PAY-0001") -> Path:
        path = self.root / "cases" / "credit-booster" / "payments" / filename
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = DEFAULT_SHEET
        worksheet.append(DEFAULT_HEADERS)
        worksheet.append(
            [
                f"{case_id} — Успешное действие",
                "Проверка подтверждённого сценария",
                "1. Подготовлены тестовые данные",
                "1. Выполнить действие",
                "1. Действие успешно выполнено",
            ]
        )
        workbook.save(path)
        return path

    def test_valid_workbook_is_indexed(self):
        self._write_case("CBN-1.xlsx")
        cases, issues = validate_repository(self.root)
        self.assertEqual([], issues)
        self.assertEqual("CB-PAY-0001", cases[0].id)
        self.assertEqual("payments", cases[0].functionality)

    def test_duplicate_id_is_rejected(self):
        self._write_case("CBN-1.xlsx")
        self._write_case("CBN-2.xlsx")
        _, issues = validate_repository(self.root)
        self.assertTrue(any("дублирующийся ID" in issue.message for issue in issues))

    def test_suite_and_export(self):
        self._write_case("CBN-1.xlsx")
        suite_path = self.root / "suites" / "payments.json"
        suite_path.write_text(
            json.dumps(
                {
                    "include": {
                        "case_ids": [],
                        "globs": [],
                        "products": [],
                        "functionalities": ["payments"],
                    },
                    "exclude": {"case_ids": [], "globs": [], "products": [], "functionalities": []},
                }
            ),
            encoding="utf-8",
        )
        cases, issues = validate_repository(self.root)
        self.assertEqual([], issues)
        selected = select_cases(cases, suite=load_suite(self.root, "payments"))
        self.assertEqual(1, len(selected))
        output = self.root / "exports" / "payments.xlsx"
        export_cases(output, selected, load_config(self.root))
        self.assertTrue(output.exists())


if __name__ == "__main__":
    unittest.main()
