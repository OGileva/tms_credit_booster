from __future__ import annotations

import csv
import fnmatch
import json
import re
import tomllib
import unicodedata
from dataclasses import asdict, dataclass
from io import StringIO
from pathlib import Path
from typing import Any, Iterable, Sequence

from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


CANONICAL_HEADERS = ["Name", "Description", "Preconditions", "Steps", "Expected Result"]
REQUIRED_SECTIONS = ["Description", "Preconditions", "Steps", "Expected Result"]
ALLOWED_STATUSES = {"active", "draft", "deprecated"}
NUMBERED_LINE = re.compile(r"^\d+\.\s+\S")
GENERIC_ID = re.compile(r"^(?P<id>[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+-\d{4})\s*(?:—|-)\s*(?P<title>.+)$")
SLUG_PATH = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*(?:/[a-z0-9]+(?:-[a-z0-9]+)*)*$")
ID_PREFIX = re.compile(r"^[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+$")
FRONT_MATTER = re.compile(r"\A\+\+\+\s*\n(?P<meta>.*?)\n\+\+\+\s*\n(?P<body>.*)\Z", re.DOTALL)


@dataclass(frozen=True)
class Section:
    product: str
    path: str
    name: str
    aliases: tuple[str, ...]
    id_prefix: str
    directory: Path


@dataclass(frozen=True)
class TestCase:
    id: str
    title: str
    description: str
    preconditions: str
    steps: str
    expected_result: str
    status: str
    tags: tuple[str, ...]
    tasks: tuple[str, ...]
    source: str | None
    product: str
    section_path: str
    section_name: str
    file: str

    @property
    def name(self) -> str:
        return f"{self.id} — {self.title}"

    def table_row(self) -> list[str]:
        return [self.name, self.description, self.preconditions, self.steps, self.expected_result]


@dataclass(frozen=True)
class Issue:
    level: str
    source: str
    message: str

    def render(self) -> str:
        return f"{self.level.upper():7} {self.source} — {self.message}"


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def load_sections(root: Path) -> list[Section]:
    cases_root = root / "cases"
    result: list[Section] = []
    for path in sorted(cases_root.rglob("section.json")) if cases_root.exists() else []:
        relative = path.parent.relative_to(cases_root)
        if len(relative.parts) < 2:
            continue
        payload = _read_json(path)
        result.append(
            Section(
                product=relative.parts[0],
                path="/".join(relative.parts[1:]),
                name=str(payload["name"]),
                aliases=tuple(str(value) for value in payload.get("aliases", [])),
                id_prefix=str(payload["id_prefix"]),
                directory=path.parent,
            )
        )
    return result


def resolve_section(root: Path, query: str, product: str | None = None) -> Section:
    query_key = query.casefold().strip()
    matches: list[Section] = []
    for section in load_sections(root):
        if product and section.product.casefold() != product.casefold():
            continue
        candidates = [section.path, section.name, *section.aliases]
        if any(candidate.casefold() == query_key for candidate in candidates):
            matches.append(section)
    if not matches:
        raise ValueError(f"раздел не найден: {query!r}; проверьте `qa-cases sections`")
    if len(matches) > 1:
        rendered = ", ".join(f"{section.product}/{section.path}" for section in matches)
        raise ValueError(f"раздел {query!r} неоднозначен: {rendered}; укажите --product")
    return matches[0]


def create_section(
    root: Path,
    *,
    product: str,
    path: str,
    name: str,
    id_prefix: str,
    aliases: Sequence[str] = (),
) -> Section:
    if not SLUG_PATH.fullmatch(path):
        raise ValueError("path должен состоять из slug-сегментов, например `scoring/loan-history`")
    if not SLUG_PATH.fullmatch(product):
        raise ValueError("product должен быть slug, например `credit-booster`")
    if not ID_PREFIX.fullmatch(id_prefix):
        raise ValueError("id-prefix должен иметь вид `CB-SCR-LH`")
    directory = root / "cases" / product / Path(path)
    config = directory / "section.json"
    if config.exists():
        raise ValueError(f"раздел уже существует: {product}/{path}")
    directory.mkdir(parents=True, exist_ok=True)
    payload = {
        "name": name.strip(),
        "aliases": sorted({alias.strip() for alias in aliases if alias.strip()}),
        "id_prefix": id_prefix,
    }
    config.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return Section(product, path, name.strip(), tuple(payload["aliases"]), id_prefix, directory)


def load_tag_dictionary(root: Path) -> tuple[set[str], dict[str, str]]:
    path = root / "dictionaries" / "tags.json"
    if not path.exists():
        return set(), {}
    payload = _read_json(path)
    known: set[str] = set()
    aliases: dict[str, str] = {}
    for item in payload.get("tags", []):
        tag_id = str(item["id"]).casefold()
        known.add(tag_id)
        aliases[tag_id] = tag_id
        for alias in item.get("aliases", []):
            aliases[str(alias).casefold()] = tag_id
    return known, aliases


def _split_tag_value(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        result: list[str] = []
        for item in value:
            result.extend(_split_tag_value(item))
        return result
    return [part.strip() for part in re.split(r"[,;\n]", str(value)) if part.strip()]


def resolve_tags(root: Path, values: Iterable[str]) -> tuple[str, ...]:
    known, aliases = load_tag_dictionary(root)
    resolved: set[str] = set()
    for raw in values:
        for value in _split_tag_value(raw):
            key = value.casefold()
            if key not in aliases:
                available = ", ".join(sorted(known))
                raise ValueError(f"неизвестный тег {value!r}; доступные теги: {available}")
            resolved.add(aliases[key])
    return tuple(sorted(resolved))


def _toml_string(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def _toml_array(values: Sequence[str]) -> str:
    return "[" + ", ".join(_toml_string(value) for value in values) + "]"


def render_case(case: TestCase) -> str:
    lines = [
        "+++",
        f"id = {_toml_string(case.id)}",
        f"title = {_toml_string(case.title)}",
        f"status = {_toml_string(case.status)}",
        f"tags = {_toml_array(case.tags)}",
        f"tasks = {_toml_array(case.tasks)}",
    ]
    if case.source:
        lines.append(f"source = {_toml_string(case.source)}")
    lines.extend(
        [
            "+++",
            "",
            f"# {case.name}",
            "",
            "## Description",
            "",
            case.description,
            "",
            "## Preconditions",
            "",
            case.preconditions,
            "",
            "## Steps",
            "",
            case.steps,
            "",
            "## Expected Result",
            "",
            case.expected_result,
            "",
        ]
    )
    return "\n".join(lines)


def write_case(path: Path, case: TestCase) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_case(case), encoding="utf-8")


def _parse_body_sections(body: str) -> dict[str, str]:
    result: dict[str, str] = {}
    matches = list(re.finditer(r"^##\s+(.+?)\s*$", body, re.MULTILINE))
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        result[match.group(1).strip()] = body[start:end].strip()
    return result


def _section_for_file(root: Path, path: Path) -> Section:
    sections = {section.directory.resolve(): section for section in load_sections(root)}
    section = sections.get(path.parent.resolve())
    if section is None:
        raise ValueError("файл кейса должен лежать непосредственно в папке с section.json")
    return section


def read_case(root: Path, path: Path) -> TestCase:
    text = path.read_text(encoding="utf-8")
    match = FRONT_MATTER.match(text)
    if not match:
        raise ValueError("нет TOML front matter между строками `+++`")
    metadata = tomllib.loads(match.group("meta"))
    body_sections = _parse_body_sections(match.group("body"))
    section = _section_for_file(root, path)
    for required in REQUIRED_SECTIONS:
        if not body_sections.get(required):
            raise ValueError(f"нет обязательного раздела `## {required}`")
    return TestCase(
        id=str(metadata.get("id", "")),
        title=str(metadata.get("title", "")),
        description=body_sections["Description"],
        preconditions=body_sections["Preconditions"],
        steps=body_sections["Steps"],
        expected_result=body_sections["Expected Result"],
        status=str(metadata.get("status", "")),
        tags=tuple(str(value) for value in metadata.get("tags", [])),
        tasks=tuple(str(value) for value in metadata.get("tasks", [])),
        source=str(metadata["source"]) if metadata.get("source") else None,
        product=section.product,
        section_path=section.path,
        section_name=section.name,
        file=path.relative_to(root).as_posix(),
    )


def case_paths(root: Path) -> list[Path]:
    cases_root = root / "cases"
    ignored = {"README.md", "AGENTS.md", "CLAUDE.md"}
    if not cases_root.exists():
        return []
    return sorted(path for path in cases_root.rglob("*.md") if path.name not in ignored)


def _numbered_list_error(value: str) -> str | None:
    lines = [line.strip() for line in value.splitlines() if line.strip()]
    if not lines:
        return "список пуст"
    if any(not NUMBERED_LINE.match(line) for line in lines):
        return "каждая строка должна начинаться с номера вида `1. `"
    numbers = [int(line.split(".", 1)[0]) for line in lines]
    if numbers != list(range(1, len(numbers) + 1)):
        return f"нумерация должна быть последовательной от 1; получено {numbers}"
    return None


def validate_repository(root: Path) -> tuple[list[TestCase], list[Issue]]:
    known_tags, _ = load_tag_dictionary(root)
    cases: list[TestCase] = []
    issues: list[Issue] = []
    for path in case_paths(root):
        source = path.relative_to(root).as_posix()
        try:
            case = read_case(root, path)
            section = _section_for_file(root, path)
        except (OSError, ValueError, tomllib.TOMLDecodeError, KeyError, TypeError) as exc:
            issues.append(Issue("error", source, str(exc)))
            continue
        if not re.fullmatch(re.escape(section.id_prefix) + r"-\d{4}", case.id):
            issues.append(Issue("error", source, f"ID должен иметь вид `{section.id_prefix}-0001`"))
        if path.stem != case.id:
            issues.append(Issue("error", source, f"имя файла должно быть `{case.id}.md`"))
        if not case.title.strip():
            issues.append(Issue("error", source, "title не заполнен"))
        if case.status not in ALLOWED_STATUSES:
            issues.append(Issue("error", source, f"status должен быть одним из {sorted(ALLOWED_STATUSES)}"))
        unknown = sorted(set(case.tags) - known_tags)
        if unknown:
            issues.append(Issue("error", source, f"неизвестные теги: {', '.join(unknown)}"))
        if list(case.tags) != sorted(set(case.tags)):
            issues.append(Issue("error", source, "теги должны быть уникальными и отсортированными"))
        for name, value in (
            ("Preconditions", case.preconditions),
            ("Steps", case.steps),
            ("Expected Result", case.expected_result),
        ):
            error = _numbered_list_error(value)
            if error:
                issues.append(Issue("error", source, f"{name}: {error}"))
        cases.append(case)

    seen_ids: dict[str, TestCase] = {}
    seen_titles: dict[tuple[str, str, str], TestCase] = {}
    for case in cases:
        if case.id in seen_ids:
            issues.append(Issue("error", case.file, f"ID дублируется с {seen_ids[case.id].file}"))
        else:
            seen_ids[case.id] = case
        key = (case.product.casefold(), case.section_path.casefold(), case.title.casefold())
        if key in seen_titles:
            issues.append(Issue("error", case.file, f"название дублируется с {seen_titles[key].file}"))
        else:
            seen_titles[key] = case
    return sorted(cases, key=lambda item: (item.product, item.section_path, item.id)), issues


def write_index(root: Path, cases: Sequence[TestCase]) -> Path:
    output = root / "generated" / "index.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = {"schema_version": 2, "case_count": len(cases), "cases": [asdict(case) for case in cases]}
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output


def _normalize_header(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return re.sub(r"[^a-zа-яё0-9]+", "", text)


HEADER_ALIASES = {
    "id": {"id", "caseid", "testcaseid", "идентификатор"},
    "name": {"name", "title", "название", "наименование"},
    "description": {"description", "описание"},
    "preconditions": {"preconditions", "precondition", "предусловия", "условия"},
    "steps": {"steps", "step", "шаги"},
    "expected_result": {"expectedresult", "expected", "ожидаемыйрезультат", "результат"},
    "tags": {"tags", "tag", "теги", "метки"},
    "status": {"status", "статус"},
    "tasks": {"tasks", "task", "requirements", "задачи", "требования"},
}


def _canonical_key(header: Any) -> str | None:
    normalized = _normalize_header(header)
    for key, aliases in HEADER_ALIASES.items():
        if normalized in aliases:
            return key
    return None


def _records_from_rows(rows: Sequence[Sequence[Any]], source: str) -> list[dict[str, Any]]:
    if not rows:
        return []
    keys = [_canonical_key(value) for value in rows[0]]
    required = {"name", "description", "preconditions", "steps", "expected_result"}
    if not required.issubset({key for key in keys if key}):
        raise ValueError(f"{source}: не найдены обязательные колонки {sorted(required)}")
    result: list[dict[str, Any]] = []
    for row in rows[1:]:
        if all(value in (None, "") for value in row):
            continue
        record: dict[str, Any] = {}
        for index, value in enumerate(row):
            if index < len(keys) and keys[index]:
                record[keys[index]] = value
        result.append(record)
    return result


def _records_from_xlsx(path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    result: list[dict[str, Any]] = []
    errors: list[str] = []
    for worksheet in workbook.worksheets:
        rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
        if not rows or all(value in (None, "") for value in rows[0]):
            continue
        try:
            result.extend(_records_from_rows(rows, f"{path.name}:{worksheet.title}"))
        except ValueError as exc:
            errors.append(str(exc))
    workbook.close()
    if not result:
        raise ValueError("; ".join(errors) if errors else f"{path.name}: тест-кейсы не найдены")
    return result


def _records_from_csv(path: Path) -> list[dict[str, Any]]:
    text = path.read_text(encoding="utf-8-sig")
    dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    return _records_from_rows(list(csv.reader(StringIO(text), dialect)), path.name)


def _records_from_json(path: Path) -> list[dict[str, Any]]:
    payload = _read_json(path)
    records = payload.get("cases", []) if isinstance(payload, dict) else payload
    if not isinstance(records, list) or not all(isinstance(item, dict) for item in records):
        raise ValueError("JSON должен содержать массив объектов или объект с массивом `cases`")
    return [{_canonical_key(key) or key: value for key, value in item.items()} for item in records]


def _records_from_docx(path: Path) -> list[dict[str, Any]]:
    document = Document(path)
    result: list[dict[str, Any]] = []
    errors: list[str] = []
    for index, table in enumerate(document.tables, start=1):
        rows = [[cell.text for cell in row.cells] for row in table.rows]
        try:
            result.extend(_records_from_rows(rows, f"{path.name}:таблица-{index}"))
        except ValueError as exc:
            errors.append(str(exc))
    if not result:
        raise ValueError("; ".join(errors) if errors else f"{path.name}: таблицы с тест-кейсами не найдены")
    return result


def _records_from_markdown(path: Path) -> list[dict[str, Any]]:
    text = path.read_text(encoding="utf-8")
    if FRONT_MATTER.match(text):
        raise ValueError("канонический Markdown-кейс следует добавлять как готовый файл после проверки ID и раздела")
    table_lines = [line.strip() for line in text.splitlines() if line.strip().startswith("|")]
    if len(table_lines) < 3:
        raise ValueError("Markdown должен содержать таблицу с обязательными колонками")
    rows: list[list[str]] = []
    for index, line in enumerate(table_lines):
        values = [value.strip().replace("<br>", "\n") for value in line.strip("|").split("|")]
        if index == 1 and all(re.fullmatch(r":?-{3,}:?", value) for value in values):
            continue
        rows.append(values)
    return _records_from_rows(rows, path.name)


def records_from_source(path: Path) -> list[dict[str, Any]]:
    suffix = path.suffix.casefold()
    if suffix == ".xlsx":
        return _records_from_xlsx(path)
    if suffix in {".csv", ".tsv"}:
        return _records_from_csv(path)
    if suffix == ".json":
        return _records_from_json(path)
    if suffix == ".docx":
        return _records_from_docx(path)
    if suffix in {".md", ".markdown"}:
        return _records_from_markdown(path)
    raise ValueError("поддерживаются XLSX, CSV/TSV, JSON, DOCX и Markdown; другой формат должен нормализовать AI-агент")


def _normalize_numbered(value: Any, field: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValueError(f"не заполнено поле {field}")
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    normalized: list[str] = []
    for index, line in enumerate(lines, start=1):
        content = re.sub(r"^(?:\d+[.)]|[-*•])\s*", "", line).strip()
        if not content:
            raise ValueError(f"{field}: пустой пункт списка")
        normalized.append(f"{index}. {content}")
    return "\n".join(normalized)


def import_cases(
    root: Path,
    *,
    source_path: Path,
    section: Section,
    command_tags: Sequence[str] = (),
    status: str = "active",
    tasks: Sequence[str] = (),
    dry_run: bool = False,
) -> list[TestCase]:
    if status not in ALLOWED_STATUSES:
        raise ValueError(f"status должен быть одним из {sorted(ALLOWED_STATUSES)}")
    existing, issues = validate_repository(root)
    errors = [issue for issue in issues if issue.level == "error"]
    if errors:
        raise ValueError("сначала исправьте текущие ошибки репозитория: " + "; ".join(issue.render() for issue in errors[:5]))
    records = records_from_source(source_path)
    if not records:
        raise ValueError("во входном файле нет тест-кейсов")

    existing_ids = {case.id for case in existing}
    existing_titles = {(case.product.casefold(), case.section_path.casefold(), case.title.casefold()) for case in existing}
    prefix_re = re.compile(rf"^{re.escape(section.id_prefix)}-(\d{{4}})$")
    used_numbers = [int(match.group(1)) for case_id in existing_ids if (match := prefix_re.match(case_id))]
    next_number = max(used_numbers, default=0) + 1
    pending: list[TestCase] = []

    for position, record in enumerate(records, start=1):
        raw_name = str(record.get("name") or "").strip()
        raw_id = str(record.get("id") or "").strip()
        title = raw_name
        name_match = GENERIC_ID.match(raw_name)
        if name_match:
            raw_id = raw_id or name_match.group("id")
            title = name_match.group("title").strip()
        if not title:
            raise ValueError(f"строка {position}: не заполнено название")
        if raw_id:
            if not prefix_re.fullmatch(raw_id):
                raise ValueError(f"строка {position}: ID должен иметь вид `{section.id_prefix}-0001`")
            case_id = raw_id
        else:
            while f"{section.id_prefix}-{next_number:04d}" in existing_ids:
                next_number += 1
            case_id = f"{section.id_prefix}-{next_number:04d}"
            next_number += 1
        if case_id in existing_ids or any(case.id == case_id for case in pending):
            raise ValueError(f"строка {position}: ID {case_id} уже существует")

        description = str(record.get("description") or "").strip()
        if not description:
            raise ValueError(f"строка {position}: не заполнено поле Description")
        tags = resolve_tags(root, [*command_tags, *_split_tag_value(record.get("tags"))])
        record_tasks = _split_tag_value(record.get("tasks"))
        case_tasks = tuple(sorted({str(value).strip() for value in [*tasks, *record_tasks] if str(value).strip()}))
        record_status = str(record.get("status") or status).casefold()
        if record_status not in ALLOWED_STATUSES:
            raise ValueError(f"строка {position}: неизвестный status {record_status!r}")
        title_key = (section.product.casefold(), section.path.casefold(), title.casefold())
        if title_key in existing_titles or any(case.title.casefold() == title.casefold() for case in pending):
            raise ValueError(f"строка {position}: кейс с названием {title!r} уже существует в разделе")
        case = TestCase(
            id=case_id,
            title=title,
            description=description,
            preconditions=_normalize_numbered(record.get("preconditions"), "Preconditions"),
            steps=_normalize_numbered(record.get("steps"), "Steps"),
            expected_result=_normalize_numbered(record.get("expected_result"), "Expected Result"),
            status=record_status,
            tags=tags,
            tasks=case_tasks,
            source=source_path.name,
            product=section.product,
            section_path=section.path,
            section_name=section.name,
            file=(section.directory / f"{case_id}.md").relative_to(root).as_posix(),
        )
        pending.append(case)
    if not dry_run:
        for case in pending:
            write_case(root / case.file, case)
    return pending


def _matches(value: str, patterns: Iterable[str]) -> bool:
    folded = value.casefold()
    return any(fnmatch.fnmatchcase(folded, pattern.casefold()) for pattern in patterns)


def load_suite(root: Path, name: str) -> dict[str, Any]:
    filename = name if name.endswith(".json") else f"{name}.json"
    path = root / "suites" / filename
    if not path.exists():
        raise ValueError(f"suite не найден: {filename}")
    return _read_json(path)


def _matches_suite_rules(case: TestCase, rules: dict[str, Any]) -> bool:
    return any(
        (
            case.id in rules.get("case_ids", []),
            _matches(case.product, rules.get("products", [])),
            _matches(case.section_path, rules.get("sections", [])),
            bool(set(case.tags) & set(rules.get("tags", []))),
            case.status in rules.get("statuses", []),
        )
    )


def _matches_section_query(case: TestCase, query: str) -> bool:
    return _matches(case.section_path, [query]) or _matches(case.section_name, [query])


def select_cases(
    cases: Sequence[TestCase],
    *,
    suite: dict[str, Any] | None = None,
    products: Sequence[str] = (),
    sections: Sequence[str] = (),
    tags: Sequence[str] = (),
    ids: Sequence[str] = (),
    search: str | None = None,
    status: str = "active",
) -> list[TestCase]:
    selected = list(cases)
    if suite is not None:
        selected = [case for case in selected if _matches_suite_rules(case, suite.get("include", {}))]
        selected = [case for case in selected if not _matches_suite_rules(case, suite.get("exclude", {}))]
    if products:
        selected = [case for case in selected if _matches(case.product, products)]
    if sections:
        selected = [case for case in selected if any(_matches_section_query(case, query) for query in sections)]
    if tags:
        wanted = {tag.casefold() for tag in tags}
        selected = [case for case in selected if wanted.issubset({tag.casefold() for tag in case.tags})]
    if ids:
        wanted_ids = {case_id.casefold() for case_id in ids}
        selected = [case for case in selected if case.id.casefold() in wanted_ids]
    if status != "all":
        selected = [case for case in selected if case.status == status]
    if search:
        needle = search.casefold()
        selected = [
            case
            for case in selected
            if needle
            in "\n".join(
                [
                    case.id,
                    case.title,
                    case.description,
                    case.preconditions,
                    case.steps,
                    case.expected_result,
                    " ".join(case.tags),
                    " ".join(case.tasks),
                    case.section_name,
                    case.section_path,
                ]
            ).casefold()
        ]
    return sorted(selected, key=lambda item: (item.product, item.section_path, item.id))


def _markdown_cell(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", "<br>")


def render_table(cases: Sequence[TestCase]) -> str:
    lines = [
        "| Name | Description | Preconditions | Steps | Expected Result |",
        "|---|---|---|---|---|",
    ]
    for case in cases:
        lines.append("| " + " | ".join(_markdown_cell(value) for value in case.table_row()) + " |")
    return "\n".join(lines) + "\n"


def export_cases(path: Path, cases: Sequence[TestCase], format_name: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if format_name in {"table", "markdown"}:
        path.write_text(render_table(cases), encoding="utf-8")
        return
    if format_name == "json":
        path.write_text(json.dumps([asdict(case) for case in cases], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return
    if format_name == "csv":
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(CANONICAL_HEADERS)
            writer.writerows(case.table_row() for case in cases)
        return
    if format_name == "xlsx":
        _export_xlsx(path, cases)
        return
    raise ValueError(f"неподдерживаемый формат: {format_name}")


def _export_xlsx(path: Path, cases: Sequence[TestCase]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Тест-кейсы"
    worksheet.append(CANONICAL_HEADERS)
    for case in cases:
        worksheet.append(case.table_row())
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
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False
    if cases:
        table = Table(displayName="TestCasesTable", ref=f"A1:E{worksheet.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        worksheet.add_table(table)
    workbook.save(path)
