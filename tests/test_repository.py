from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from qa_cases.core import (
    TestCase,
    export_cases,
    import_cases,
    load_sections,
    render_table,
    resolve_section,
    resolve_tags,
    select_cases,
    validate_repository,
    write_case,
)


class RepositoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        section = self.root / "cases" / "credit-booster" / "scoring" / "loan-history"
        section.mkdir(parents=True)
        (section / "section.json").write_text(
            json.dumps(
                {
                    "name": "История займов",
                    "aliases": ["История кредитов"],
                    "id_prefix": "CB-SCR-LH",
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        dictionaries = self.root / "dictionaries"
        dictionaries.mkdir()
        (dictionaries / "tags.json").write_text(
            json.dumps(
                {
                    "tags": [
                        {"id": "regression", "aliases": ["регресс"]},
                        {"id": "positive", "aliases": ["позитивный"]},
                    ]
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        (self.root / "generated").mkdir()

    def tearDown(self):
        self.temp.cleanup()

    def _case(self, case_id: str = "CB-SCR-LH-0001", title: str = "Отображение истории") -> TestCase:
        return TestCase(
            id=case_id,
            title=title,
            description="Проверка отображения истории займов",
            preconditions="1. Создан клиент с займом",
            steps="1. Открыть историю займов",
            expected_result="1. История займов отображена",
            status="active",
            tags=("regression",),
            tasks=("CBN-1",),
            source="cases.csv",
            product="credit-booster",
            section_path="scoring/loan-history",
            section_name="История займов",
            file=f"cases/credit-booster/scoring/loan-history/{case_id}.md",
        )

    def test_markdown_case_is_valid_and_indexed(self):
        case = self._case()
        write_case(self.root / case.file, case)
        cases, issues = validate_repository(self.root)
        self.assertEqual([], issues)
        self.assertEqual("CB-SCR-LH-0001", cases[0].id)
        self.assertEqual("История займов", cases[0].section_name)

    def test_mass_csv_import_allocates_ids_and_resolves_tag_alias(self):
        csv_path = self.root / "input.csv"
        with csv_path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["Name", "Description", "Preconditions", "Steps", "Expected Result", "Tags"])
            writer.writerow(["Первый кейс", "Описание 1", "Клиент создан", "Открыть раздел", "Раздел открыт", "регресс"])
            writer.writerow(["Второй кейс", "Описание 2", "Займ создан", "Открыть историю", "История открыта", "позитивный"])
        section = resolve_section(self.root, "История займов", "credit-booster")
        imported = import_cases(
            self.root,
            source_path=csv_path,
            section=section,
            command_tags=("regression",),
            tasks=("CBN-2",),
        )
        self.assertEqual(["CB-SCR-LH-0001", "CB-SCR-LH-0002"], [case.id for case in imported])
        self.assertTrue((section.directory / "CB-SCR-LH-0002.md").exists())
        cases, issues = validate_repository(self.root)
        self.assertEqual([], issues)
        self.assertEqual(2, len(cases))
        self.assertEqual(("positive", "regression"), cases[1].tags)

    def test_selection_by_russian_section_and_tag(self):
        case = self._case()
        selected = select_cases(
            [case],
            sections=("История займов",),
            tags=resolve_tags(self.root, ["регресс"]),
        )
        self.assertEqual([case], selected)
        table = render_table(selected)
        self.assertIn("CB-SCR-LH-0001", table)
        self.assertIn("Expected Result", table)

    def test_xlsx_export_keeps_five_canonical_columns(self):
        output = self.root / "exports" / "selection.xlsx"
        export_cases(output, [self._case()], "xlsx")
        workbook = load_workbook(output, read_only=False, data_only=False)
        worksheet = workbook["Тест-кейсы"]
        self.assertEqual(
            ["Name", "Description", "Preconditions", "Steps", "Expected Result"],
            [worksheet.cell(1, column).value for column in range(1, 6)],
        )
        self.assertEqual("CB-SCR-LH-0001 — Отображение истории", worksheet["A2"].value)
        workbook.close()

    def test_duplicate_id_is_rejected(self):
        first = self._case()
        second = self._case(title="Другое название")
        write_case(self.root / first.file, first)
        duplicate_path = self.root / "cases" / "credit-booster" / "scoring" / "loan-history" / "copy.md"
        write_case(duplicate_path, second)
        _, issues = validate_repository(self.root)
        self.assertTrue(any("имя файла" in issue.message or "ID дублируется" in issue.message for issue in issues))


if __name__ == "__main__":
    unittest.main()
