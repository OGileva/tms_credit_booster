from __future__ import annotations

import fnmatch
import json
import re
import shutil
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


DEFAULT_HEADERS = [
    "Name",
    "Description",
    "Preconditions",
    "Steps",
    "Expected Result",
]
DEFAULT_SHEET = "Тест-кейсы"
DEFAULT_ID_PATTERN = r"^(?P<id>[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+-\d{4})\s+—\s+\S"
NUMBERED_LINE = re.compile(r"^\d+\.\s+\S")
SLUG = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
PACK_NAME = re.compile(r"^[A-Za-z0-9А-Яа-яЁё._-]+$")


@dataclass(frozen=True)
class RepositoryConfig:
    sheet_name: str
    headers: tuple[str, ...]
    id_pattern: str


@dataclass(frozen=True)
class TestCase:
    id: str
    name: str
    description: str
    preconditions: str
    steps: str
    expected_result: str
    product: str
    functionality: str
    source: str
    row: int

    def excel_row(self) -> list[str]:
        return [
            self.name,
            self.description,
            self.preconditions,
            self.steps,
            self.expected_result,
        ]


@dataclass(frozen=True)
class Issue:
    level: str
    source: str
    message: str
    row: int | None = None

    def render(self) -> str:
        location = self.source if self.row is None else f"{self.source}:{self.row}"
        return f"{self.level.upper():7} {location} — {self.message}"


def load_config(root: Path) -> RepositoryConfig:
    path = root / "project.json"
    if not path.exists():
        return RepositoryConfig(DEFAULT_SHEET, tuple(DEFAULT_HEADERS), DEFAULT_ID_PATTERN)
    payload = json.loads(path.read_text(encoding="utf-8"))
    return RepositoryConfig(
        sheet_name=str(payload.get("sheet_name", DEFAULT_SHEET)),
        headers=tuple(payload.get("headers", DEFAULT_HEADERS)),
        id_pattern=str(payload.get("id_pattern", DEFAULT_ID_PATTERN)),
    )


def workbook_paths(root: Path) -> list[Path]:
    cases_root = root / "cases"
    if not cases_root.exists():
        return []
    return sorted(
        path
        for path in cases_root.rglob("*.xlsx")
        if not path.name.startswith("~$") and path.is_file()
    )


def _relative_source(root: Path, path: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


def _path_dimensions(root: Path, path: Path) -> tuple[str, str] | None:
    relative = path.resolve().relative_to((root / "cases").resolve())
    if len(relative.parts) < 3:
        return None
    product = relative.parts[0]
    functionality = "/".join(relative.parts[1:-1])
    return product, functionality


def _numbered_list_error(value: str) -> str | None:
    lines = [line.strip() for line in value.splitlines() if line.strip()]
    if not lines:
        return "список пуст"
    invalid = [line for line in lines if not NUMBERED_LINE.match(line)]
    if invalid:
        return "каждая непустая строка должна начинаться с номера вида `1. `"
    numbers = [int(line.split(".", 1)[0]) for line in lines]
    expected = list(range(1, len(numbers) + 1))
    if numbers != expected:
        return f"нумерация должна быть последовательной от 1; получено {numbers}"
    return None


def read_workbook(root: Path, path: Path, config: RepositoryConfig) -> tuple[list[TestCase], list[Issue]]:
    source = _relative_source(root, path)
    issues: list[Issue] = []
    dimensions = _path_dimensions(root, path)
    if dimensions is None:
        return [], [Issue("error", source, "ожидается путь cases/<product>/<function>/<file>.xlsx")]
    product, functionality = dimensions

    try:
        workbook = load_workbook(path, data_only=False, read_only=False)
    except Exception as exc:  # pragma: no cover - library supplies many exception types
        return [], [Issue("error", source, f"не удалось открыть XLSX: {exc}")]

    if workbook.sheetnames != [config.sheet_name]:
        issues.append(
            Issue(
                "error",
                source,
                f"ожидается единственный лист `{config.sheet_name}`; найдено: {workbook.sheetnames}",
            )
        )
    if config.sheet_name not in workbook.sheetnames:
        workbook.close()
        return [], issues

    worksheet = workbook[config.sheet_name]
    actual_headers = [worksheet.cell(1, column).value for column in range(1, len(config.headers) + 1)]
    if actual_headers != list(config.headers):
        issues.append(
            Issue(
                "error",
                source,
                f"заголовки должны быть {list(config.headers)}; получено {actual_headers}",
                1,
            )
        )

    for column in range(len(config.headers) + 1, worksheet.max_column + 1):
        if any(worksheet.cell(row, column).value not in (None, "") for row in range(1, worksheet.max_row + 1)):
            issues.append(Issue("error", source, f"обнаружены данные за пределами пяти колонок, колонка {column}"))

    id_regex = re.compile(config.id_pattern)
    cases: list[TestCase] = []
    for row_number in range(2, worksheet.max_row + 1):
        cells = [worksheet.cell(row_number, column) for column in range(1, len(config.headers) + 1)]
        values = [cell.value for cell in cells]
        if all(value in (None, "") for value in values):
            continue
        for cell in cells:
            if cell.data_type == "f":
                issues.append(Issue("error", source, "формулы в тест-кейсах запрещены", row_number))
        missing = [config.headers[index] for index, value in enumerate(values) if value in (None, "")]
        if missing:
            issues.append(Issue("error", source, f"не заполнены поля: {', '.join(missing)}", row_number))
            continue
        normalized = [str(value).strip() for value in values]
        name, description, preconditions, steps, expected_result = normalized
        id_match = id_regex.match(name)
        if not id_match:
            issues.append(
                Issue(
                    "error",
                    source,
                    "Name должен начинаться со стабильного ID и тире, например `CB-PAY-0001 — ...`",
                    row_number,
                )
            )
            continue
        for field_name, value in (
            ("Preconditions", preconditions),
            ("Steps", steps),
            ("Expected Result", expected_result),
        ):
            list_error = _numbered_list_error(value)
            if list_error:
                issues.append(Issue("error", source, f"{field_name}: {list_error}", row_number))
        cases.append(
            TestCase(
                id=id_match.group("id"),
                name=name,
                description=description,
                preconditions=preconditions,
                steps=steps,
                expected_result=expected_result,
                product=product,
                functionality=functionality,
                source=source,
                row=row_number,
            )
        )
    workbook.close()
    return cases, issues


def validate_repository(root: Path) -> tuple[list[TestCase], list[Issue]]:
    config = load_config(root)
    all_cases: list[TestCase] = []
    issues: list[Issue] = []
    for path in workbook_paths(root):
        cases, workbook_issues = read_workbook(root, path, config)
        all_cases.extend(cases)
        issues.extend(workbook_issues)

    seen_ids: dict[str, TestCase] = {}
    seen_names: dict[str, TestCase] = {}
    for case in all_cases:
        if case.id in seen_ids:
            first = seen_ids[case.id]
            issues.append(
                Issue(
                    "error",
                    case.source,
                    f"дублирующийся ID {case.id}; первый найден в {first.source}:{first.row}",
                    case.row,
                )
            )
        else:
            seen_ids[case.id] = case
        name_key = case.name.casefold()
        if name_key in seen_names:
            first = seen_names[name_key]
            issues.append(
                Issue(
                    "error",
                    case.source,
                    f"дублирующийся Name; первый найден в {first.source}:{first.row}",
                    case.row,
                )
            )
        else:
            seen_names[name_key] = case
    return sorted(all_cases, key=lambda case: (case.product, case.functionality, case.id)), issues


def write_index(root: Path, cases: Sequence[TestCase]) -> Path:
    output = root / "generated" / "index.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "schema_version": 1,
        "case_count": len(cases),
        "cases": [asdict(case) for case in cases],
    }
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output


def load_suite(root: Path, name: str) -> dict[str, Any]:
    filename = name if name.endswith(".json") else f"{name}.json"
    path = root / "suites" / filename
    if not path.exists():
        raise ValueError(f"suite не найден: {path.relative_to(root)}")
    payload = json.loads(path.read_text(encoding="utf-8"))
    for section in ("include", "exclude"):
        rules = payload.get(section, {})
        if not isinstance(rules, dict):
            raise ValueError(f"{filename}: `{section}` должен быть объектом")
        for key in ("case_ids", "globs", "products", "functionalities"):
            values = rules.get(key, [])
            if not isinstance(values, list) or not all(isinstance(value, str) for value in values):
                raise ValueError(f"{filename}: `{section}.{key}` должен быть списком строк")
    return payload


def _matches_patterns(value: str, patterns: Iterable[str]) -> bool:
    folded = value.casefold()
    return any(fnmatch.fnmatchcase(folded, pattern.casefold()) for pattern in patterns)


def _matches_rules(case: TestCase, rules: dict[str, Any]) -> bool:
    return any(
        (
            case.id in rules.get("case_ids", []),
            _matches_patterns(case.source, rules.get("globs", [])),
            _matches_patterns(case.product, rules.get("products", [])),
            _matches_patterns(case.functionality, rules.get("functionalities", [])),
        )
    )


def select_cases(
    cases: Sequence[TestCase],
    *,
    suite: dict[str, Any] | None = None,
    products: Sequence[str] = (),
    functionalities: Sequence[str] = (),
    ids: Sequence[str] = (),
    search: str | None = None,
) -> list[TestCase]:
    selected = list(cases)
    if suite is not None:
        selected = [case for case in selected if _matches_rules(case, suite.get("include", {}))]
        selected = [case for case in selected if not _matches_rules(case, suite.get("exclude", {}))]
    if products:
        selected = [case for case in selected if _matches_patterns(case.product, products)]
    if functionalities:
        selected = [case for case in selected if _matches_patterns(case.functionality, functionalities)]
    if ids:
        wanted = {value.casefold() for value in ids}
        selected = [case for case in selected if case.id.casefold() in wanted]
    if search:
        needle = search.casefold()
        selected = [
            case
            for case in selected
            if needle
            in "\n".join(
                [
                    case.id,
                    case.name,
                    case.description,
                    case.preconditions,
                    case.steps,
                    case.expected_result,
                    case.product,
                    case.functionality,
                ]
            ).casefold()
        ]
    return sorted(selected, key=lambda case: (case.product, case.functionality, case.id))


def export_cases(path: Path, cases: Sequence[TestCase], config: RepositoryConfig) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = config.sheet_name
    worksheet.append(list(config.headers))
    for case in cases:
        worksheet.append(case.excel_row())

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10, color="222222")
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for column, width in zip("ABCDE", (42, 54, 54, 54, 54), strict=True):
        worksheet.column_dimensions[column].width = width
    worksheet.row_dimensions[1].height = 28
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False
    worksheet.auto_filter.ref = f"A1:E{max(1, worksheet.max_row)}"

    if cases:
        table = Table(displayName="TestCasesTable", ref=f"A1:E{worksheet.max_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def create_pack(root: Path, product: str, functionality: str, name: str) -> Path:
    if not SLUG.fullmatch(product):
        raise ValueError("product должен быть slug в нижнем регистре, например `credit-booster`")
    if not SLUG.fullmatch(functionality):
        raise ValueError("function должен быть slug в нижнем регистре, например `client-card`")
    if not PACK_NAME.fullmatch(name):
        raise ValueError("name содержит недопустимые символы")
    template = root / "_templates" / "test-cases.xlsx"
    if not template.exists():
        raise ValueError(f"шаблон не найден: {template.relative_to(root)}")
    output = root / "cases" / product / functionality / f"{name}.xlsx"
    if output.exists():
        raise ValueError(f"файл уже существует: {output.relative_to(root)}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    return output
