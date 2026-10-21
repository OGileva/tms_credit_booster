from __future__ import annotations

import csv
import fnmatch
import html
import json
import re
import tomllib
import unicodedata
from dataclasses import asdict, dataclass
from io import StringIO
from pathlib import Path
from typing import Any, Iterable, Sequence

from docx import Document
from jsonschema import Draft202012Validator, SchemaError
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


CANONICAL_HEADERS = ["Name", "Description", "Preconditions", "Steps", "Expected Result"]
METADATA_HEADERS = ["ID", "Tags", "Status", "Tasks", "Criticality"]
CRITICALITY_VALUES = {"critical", "medium", "low"}
REQUIRED_SECTIONS = ["Preconditions", "Steps", "Expected Result"]
ALLOWED_STATUSES = {"active", "draft", "deprecated"}
NUMBERED_LINE = re.compile(r"^\d+\.\s+\S")
GENERIC_ID = re.compile(r"^(?P<id>(?:[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+-\d{1,6}|[1-9]\d{0,5}))\s*(?:—|-)\s*(?P<title>.+)$")
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
    criticality: str | None = None

    @property
    def name(self) -> str:
        return f"{self.id} — {self.title}"

    def table_row(self, *, include_metadata: bool = False) -> list[str]:
        row = [self.name, self.description, self.preconditions, self.steps, self.expected_result]
        if include_metadata:
            row.extend([self.id, ", ".join(self.tags), self.status, ", ".join(self.tasks), self.criticality or ""])
        return row


@dataclass(frozen=True)
class Issue:
    level: str
    source: str
    message: str

    def render(self) -> str:
        return f"{self.level.upper():7} {self.source} — {self.message}"


def _read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ValueError(f"{path}: {exc}") from exc


def _validate_document(root: Path, schema_name: str, payload: Any, source: str) -> None:
    schema = _read_json(root / "schemas" / f"{schema_name}.schema.json")
    try:
        Draft202012Validator.check_schema(schema)
    except SchemaError as exc:
        raise ValueError(f"некорректная схема {schema_name}: {exc.message}") from exc
    errors = list(Draft202012Validator(schema).iter_errors(payload))
    if errors:
        details = "; ".join(
            f"{'.'.join(str(part) for part in error.absolute_path) or '<root>'}: {error.message}"
            for error in errors
        )
        raise ValueError(f"{source}: {details}")


def load_sections(root: Path) -> list[Section]:
    cases_root = root / "cases"
    result: list[Section] = []
    prefixes: dict[str, Path] = {}
    for path in sorted(cases_root.rglob("section.json")) if cases_root.exists() else []:
        relative = path.parent.relative_to(cases_root)
        if len(relative.parts) < 2:
            raise ValueError(f"{path}: раздел должен находиться внутри папки продукта")
        if any(not SLUG_PATH.fullmatch(part) for part in relative.parts):
            raise ValueError(f"{path}: путь раздела должен состоять из slug-сегментов")
        payload = _read_json(path)
        _validate_document(root, "section", payload, str(path))
        prefix = payload["id_prefix"]
        if prefix in prefixes:
            raise ValueError(f"{path}: id_prefix {prefix} дублируется с {prefixes[prefix]}")
        prefixes[prefix] = path
        result.append(
            Section(
                product=relative.parts[0],
                path="/".join(relative.parts[1:]),
                name=payload["name"],
                aliases=tuple(payload.get("aliases", [])),
                id_prefix=prefix,
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
        raise ValueError("path должен состоять из slug-сегментов, например `module/feature`")
    if not SLUG_PATH.fullmatch(product) or "/" in product:
        raise ValueError("product должен быть slug, например `credit-booster`")
    if not ID_PREFIX.fullmatch(id_prefix):
        raise ValueError("id-prefix должен иметь вид `TEST-SECTION`")
    directory = root / "cases" / product / Path(path)
    config = directory / "section.json"
    if config.exists():
        raise ValueError(f"раздел уже существует: {product}/{path}")
    for section in load_sections(root):
        if section.id_prefix == id_prefix:
            raise ValueError(f"id_prefix {id_prefix} уже используется в {section.product}/{section.path}")
    payload = {
        "name": name.strip(),
        "aliases": sorted({alias.strip() for alias in aliases if alias.strip()}),
        "id_prefix": id_prefix,
    }
    _validate_document(root, "section", payload, str(config))
    directory.mkdir(parents=True, exist_ok=True)
    config.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return Section(product, path, name.strip(), tuple(payload["aliases"]), id_prefix, directory)


def load_tag_dictionary(root: Path) -> tuple[set[str], dict[str, str]]:
    path = root / "dictionaries" / "tags.json"
    payload = _read_json(path)
    _validate_document(root, "tags", payload, str(path))
    known: set[str] = set()
    aliases: dict[str, str] = {}
    for item in payload["tags"]:
        tag_id = item["id"]
        if tag_id in known:
            raise ValueError(f"{path}: ID тега {tag_id!r} дублируется")
        known.add(tag_id)
        aliases[tag_id] = tag_id
    for item in payload["tags"]:
        tag_id = item["id"]
        for alias in item.get("aliases", []):
            key = alias.casefold().strip()
            if key in aliases and aliases[key] != tag_id:
                raise ValueError(f"{path}: alias тега {alias!r} неоднозначен")
            aliases[key] = tag_id
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
    if case.criticality:
        lines.append(f"criticality = {_toml_string(case.criticality)}")
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


def grouped_case_path(root: Path, section: Section) -> Path:
    label = next((alias for alias in section.aliases if "/" in alias), section.name)
    parts = [part.strip().casefold() for part in label.split("/") if part.strip()]
    # Нормализуем только названия верхнего уровня, для которых в правилах
    # явно задана форма единственного числа. Окончание «и» не является
    # универсальным признаком множественного числа: «заявки» должно
    # сохраняться без усечения.
    singular_aliases = {"листинги": "листинг", "допуслуги": "допуслуга"}
    parts = [singular_aliases.get(part, part) for part in parts]
    filename = " ".join(dict.fromkeys((parts[0], parts[-1]))) if parts else section.name.casefold()
    return section.directory / f"{filename}.md"


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
    _validate_document(root, "test-case", metadata, path.relative_to(root).as_posix())
    body_sections = _parse_body_sections(match.group("body"))
    section = _section_for_file(root, path)
    for required in REQUIRED_SECTIONS:
        if not body_sections.get(required):
            raise ValueError(f"нет обязательного раздела `## {required}`")
    return TestCase(
        id=metadata["id"],
        title=metadata["title"],
        description=body_sections["Description"],
        preconditions=body_sections["Preconditions"],
        steps=body_sections["Steps"],
        expected_result=body_sections["Expected Result"],
        status=metadata["status"],
        tags=tuple(metadata["tags"]),
        tasks=tuple(metadata["tasks"]),
        source=metadata.get("source"),
        product=section.product,
        section_path=section.path,
        section_name=section.name,
        file=path.relative_to(root).as_posix(),
        criticality=metadata.get("criticality"),
    )


def read_grouped_cases(root: Path, path: Path) -> list[TestCase]:
    text = path.read_text(encoding="utf-8")
    chunks = re.findall(r"\+\+\+\s*\n(.*?)\n\+\+\+\s*\n(.*?)(?=\n\+\+\+\s*\n|\Z)", text, re.DOTALL)
    if not chunks:
        raise ValueError("групповой файл не содержит блоков кейсов")
    result = []
    for meta, body in chunks:
        temporary = path.parent / f"{tomllib.loads(meta)['id']}.md"
        temporary_text = f"+++\n{meta}\n+++\n{body}"
        # Parse the same canonical block without creating a temporary file.
        metadata = tomllib.loads(meta)
        _validate_document(root, "test-case", metadata, path.relative_to(root).as_posix())
        sections = _parse_body_sections(body)
        section = _section_for_file(root, path)
        for required in REQUIRED_SECTIONS:
            if not sections.get(required):
                raise ValueError(f"нет обязательного раздела `## {required}`")
        result.append(TestCase(metadata["id"], metadata["title"], sections.get("Description", ""),
                               sections["Preconditions"], sections["Steps"], sections["Expected Result"],
                               metadata["status"], tuple(metadata["tags"]), tuple(metadata["tasks"]),
                               metadata.get("source"), section.product, section.path, section.name,
                               path.relative_to(root).as_posix(), metadata.get("criticality")))
    return result


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
    cases: list[TestCase] = []
    issues: list[Issue] = []
    known_tags: set[str] = set()
    try:
        known_tags, _ = load_tag_dictionary(root)
    except (OSError, ValueError, KeyError, TypeError) as exc:
        issues.append(Issue("error", "dictionaries/tags.json", str(exc)))
    try:
        load_sections(root)
    except (OSError, ValueError, KeyError, TypeError) as exc:
        issues.append(Issue("error", "cases", str(exc)))
    for path in case_paths(root):
        source = path.relative_to(root).as_posix()
        try:
            if re.fullmatch(r"[1-9]\d{0,5}", path.stem):
                parsed_cases = [read_case(root, path)]
            else:
                try:
                    parsed_cases = read_grouped_cases(root, path)
                except ValueError as grouped_error:
                    if "два или более" not in str(grouped_error):
                        raise
                    parsed_cases = [read_case(root, path)]
        except (OSError, ValueError, tomllib.TOMLDecodeError, KeyError, TypeError) as exc:
            issues.append(Issue("error", source, str(exc)))
            continue
        for case in parsed_cases:
            if not re.fullmatch(r"[1-9]\d{0,5}", case.id):
                issues.append(Issue("error", source, "ID должен быть числом от 1 до 999999"))
            if len(parsed_cases) == 1 and re.fullmatch(r"[1-9]\d{0,5}", path.stem) and path.stem != case.id:
                issues.append(Issue("error", source, f"имя файла должно быть `{case.id}.md`"))
            if not case.title.strip():
                issues.append(Issue("error", source, "title не заполнен"))
            if case.status not in ALLOWED_STATUSES:
                issues.append(Issue("error", source, f"status должен быть одним из {sorted(ALLOWED_STATUSES)}"))
            unknown = sorted(set(case.tags) - known_tags)
            if unknown:
                issues.append(Issue("error", source, f"неизвестные теги: {', '.join(unknown)}"))
            for name, value in (("Preconditions", case.preconditions), ("Steps", case.steps), ("Expected Result", case.expected_result)):
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


def _normalize_header(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return re.sub(r"[^a-zа-яё0-9]+", "", text)


HEADER_ALIASES = {
    "id": {"id", "caseid", "testcaseid", "идентификатор"},
    "name": {"name", "title", "название", "наименование"},
    "description": {"description", "описание"},
    "preconditions": {"preconditions", "precondition", "предусловия", "условия"},
    "steps": {"steps", "step", "шаги", "шагивоспроизведения"},
    "expected_result": {"expectedresult", "expected", "ожидаемыйрезультат", "результат"},
    "tags": {"tags", "tag", "теги", "метки", "тег"},
    "criticality": {"criticality", "критичность", "критичностьтеста"},
    "status": {"status", "статус"},
    "tasks": {"tasks", "task", "requirements", "задачи", "требования"},
}


def _canonical_key(header: Any) -> str | None:
    normalized = _normalize_header(header)
    for key, aliases in HEADER_ALIASES.items():
        if normalized in aliases:
            return key
    return None


class ImportRecord(dict[str, Any]):
    def __init__(self, values: dict[str, Any], location: str):
        super().__init__(values)
        self.location = location


def _records_from_rows(rows: Sequence[Sequence[Any]], source: str) -> list[dict[str, Any]]:
    header_index = next(
        (index for index, row in enumerate(rows) if any(value not in (None, "") for value in row)), None
    )
    if header_index is None:
        return []
    keys = [_canonical_key(value) for value in rows[header_index]]
    required = {"name", "preconditions", "steps", "expected_result"}
    if not required.issubset({key for key in keys if key}):
        raise ValueError(f"{source}: не найдены обязательные колонки {sorted(required)}")
    known_keys = [key for key in keys if key]
    if len(known_keys) != len(set(known_keys)):
        raise ValueError(f"{source}: обязательные или служебные колонки дублируются")
    result: list[dict[str, Any]] = []
    for position, row in enumerate(rows[header_index + 1:], start=header_index + 2):
        if all(value in (None, "") for value in row):
            continue
        if any(value not in (None, "") for value in row[len(keys):]):
            raise ValueError(f"{source}:строка {position}: значения выходят за пределы колонок")
        record: dict[str, Any] = {}
        for index, value in enumerate(row):
            if index < len(keys) and keys[index]:
                record[keys[index]] = value
        result.append(ImportRecord(record, f"{source}:строка {position}"))
    return result


def _records_from_xlsx(path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    result: list[dict[str, Any]] = []
    errors: list[str] = []
    try:
        for worksheet in workbook.worksheets:
            rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
            try:
                result.extend(_records_from_rows(rows, f"{path.name}:{worksheet.title}"))
            except ValueError as exc:
                errors.append(str(exc))
    finally:
        workbook.close()
    if errors:
        raise ValueError("; ".join(errors))
    if not result:
        raise ValueError(f"{path.name}: тест-кейсы не найдены")
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
    return [
        ImportRecord({_canonical_key(key) or key: value for key, value in item.items()}, f"{path.name}:строка {index}")
        for index, item in enumerate(records, start=1)
    ]


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
    if errors:
        raise ValueError("; ".join(errors))
    if not result:
        raise ValueError(f"{path.name}: таблицы с тест-кейсами не найдены")
    return result


def _markdown_row(line: str) -> list[str]:
    # Split only on unescaped pipes; decode our escaped backslashes and pipes.
    cells: list[str] = []
    current: list[str] = []
    index = 1  # All accepted table rows begin with an outer pipe.
    while index < len(line):
        char = line[index]
        if char == "\\" and index + 1 < len(line) and line[index + 1] in "\\|":
            current.append(line[index + 1])
            index += 2
            continue
        if char == "|":
            cells.append("".join(current).strip())
            current = []
        else:
            current.append(char)
        index += 1
    if current or not line.endswith("|"):
        cells.append("".join(current).strip())
    return [html.unescape(re.sub(r"<br\s*/?>", "\n", value, flags=re.IGNORECASE)) for value in cells]


def _records_from_markdown(path: Path) -> list[dict[str, Any]]:
    text = path.read_text(encoding="utf-8")
    if FRONT_MATTER.match(text):
        raise ValueError("канонический Markdown-кейс следует добавлять как готовый файл после проверки ID и раздела")
    table_lines = [line.strip() for line in text.splitlines() if line.strip().startswith("|")]
    if len(table_lines) < 3:
        raise ValueError("Markdown должен содержать таблицу с обязательными колонками")
    rows: list[list[str]] = []
    for index, line in enumerate(table_lines):
        values = _markdown_row(line)
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
    registry_path = root / "id-registry.json"
    if registry_path.exists():
        registry = _read_json(registry_path)
    else:
        registry = {"version": 1, "assigned": {}}
    assigned = set(registry.get("assigned", {}).get(f"{section.product}/{section.path}", []))
    existing_titles = {(case.product.casefold(), case.section_path.casefold(), case.title.casefold()) for case in existing}
    existing_contents = {
        (case.product.casefold(), case.section_path.casefold(), case.title, case.preconditions, case.steps, case.expected_result)
        for case in existing
    }
    prefix_re = re.compile(r"^(\d{1,6})$")
    used_numbers = [int(match.group(1)) for case_id in (*existing_ids, *assigned) if (match := prefix_re.match(case_id))]
    next_number = max(used_numbers, default=0) + 1
    pending: list[TestCase] = []
    prepared: list[tuple[dict[str, Any], str, str]] = []
    for position, record in enumerate(records, start=1):
        location = getattr(record, "location", f"{source_path.name}:строка {position}")
        raw_name = str(record.get("name") or "").strip()
        title = raw_name
        name_match = GENERIC_ID.match(raw_name)
        if name_match:
            title = name_match.group("title").strip()
        if not title:
            raise ValueError(f"{location}: не заполнено название")
        prepared.append((record, location, title))

    pending_contents: set[tuple[str, str, str, str, str, str]] = set()
    for record, location, title in prepared:
        description = str(record.get("description") or "").strip()
        try:
            tags = resolve_tags(root, [*command_tags, *_split_tag_value(record.get("tags"))])
            preconditions = _normalize_numbered(record.get("preconditions"), "Preconditions")
            steps = _normalize_numbered(record.get("steps"), "Steps")
            expected_result = _normalize_numbered(record.get("expected_result"), "Expected Result")
        except ValueError as exc:
            raise ValueError(f"{location}: {exc}") from exc
        record_tasks = _split_tag_value(record.get("tasks"))
        content_key = (section.product.casefold(), section.path.casefold(), title, preconditions, steps, expected_result)
        if content_key in existing_contents or content_key in pending_contents:
            continue
        pending_contents.add(content_key)
        while next_number <= 999999 and str(next_number) in existing_ids:
            next_number += 1
        if next_number > 999999:
            raise ValueError(f"{location}: диапазон ID 1…999999 исчерпан")
        case_id = str(next_number)
        existing_ids.add(case_id)
        next_number += 1
        case_tasks = tuple(sorted({str(value).strip() for value in [*tasks, *record_tasks] if str(value).strip()}))
        record_status = str(record.get("status") or status).casefold()
        if record_status not in ALLOWED_STATUSES:
            raise ValueError(f"{location}: неизвестный status {record_status!r}")
        criticality = str(record.get("criticality") or "").casefold().strip() or None
        if criticality and criticality not in CRITICALITY_VALUES:
            raise ValueError(f"{location}: критичность должна быть одной из critical, medium, low")
        title_key = (section.product.casefold(), section.path.casefold(), title.casefold())
        if title_key in existing_titles or any(case.title.casefold() == title.casefold() for case in pending):
            raise ValueError(f"{location}: кейс с названием {title!r} уже существует в разделе")
        case = TestCase(
            id=case_id,
            title=title,
            description=description,
            preconditions=preconditions,
            steps=steps,
            expected_result=expected_result,
            status=record_status,
            tags=tags,
            tasks=case_tasks,
            source=source_path.name,
            product=section.product,
            section_path=section.path,
            section_name=section.name,
            file=(section.directory / f"{case_id}.md").relative_to(root).as_posix(),
            criticality=criticality,
        )
        pending.append(case)
    if not dry_run:
        if section.product == "test-product":
            for case in pending:
                write_case(root / case.file, case)
        else:
            group_path = grouped_case_path(root, section)
            existing_text = group_path.read_text(encoding="utf-8") if group_path.exists() else "# " + group_path.stem.title() + "\n\n"
            group_path.write_text(existing_text.rstrip() + "\n\n" + "\n\n".join(render_case(case) for case in pending) + "\n", encoding="utf-8")
        key = f"{section.product}/{section.path}"
        registry.setdefault("assigned", {}).setdefault(key, [])
        registry["assigned"][key] = sorted(set(registry["assigned"][key]) | {case.id for case in pending})
        registry_path.write_text(json.dumps(registry, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return pending


def _matches(value: str, patterns: Iterable[str]) -> bool:
    folded = value.casefold()
    return any(fnmatch.fnmatchcase(folded, pattern.casefold()) for pattern in patterns)


def resolve_section_queries(root: Path, queries: Sequence[str], products: Sequence[str] = ()) -> tuple[str, ...]:
    """Resolve exact names/aliases once, retaining glob and multi-product filters."""
    available = [section for section in load_sections(root) if not products or _matches(section.product, products)]
    resolved: list[str] = []
    for query in queries:
        query_key = query.casefold().strip()
        matches = [
            section for section in available
            if any(value.casefold() == query_key for value in (section.path, section.name, *section.aliases))
        ]
        if len(matches) > 1:
            options = ", ".join(f"{section.product}/{section.path}" for section in matches)
            raise ValueError(f"раздел {query!r} неоднозначен: {options}; уточните --product")
        if matches:
            section = matches[0]
            resolved.append(f"{section.product}/{section.path}")
        elif any(char in query for char in "*?["):
            resolved.append(query)
        else:
            raise ValueError(f"раздел не найден: {query!r}; проверьте `qa-cases sections`")
    return tuple(resolved)


def _matches_section_query(case: TestCase, query: str) -> bool:
    candidates = (case.section_path, case.section_name, f"{case.product}/{case.section_path}")
    return any(_matches(value, [query]) for value in candidates)


def select_cases(
    cases: Sequence[TestCase],
    *,
    products: Sequence[str] = (),
    sections: Sequence[str] = (),
    tags: Sequence[str] = (),
    ids: Sequence[str] = (),
    search: str | None = None,
    status: str = "active",
) -> list[TestCase]:
    selected = list(cases)
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
    return html.escape(value, quote=False).replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def render_table(cases: Sequence[TestCase], *, include_metadata: bool = False) -> str:
    headers = CANONICAL_HEADERS + (METADATA_HEADERS if include_metadata else [])
    lines = [
        "| " + " | ".join(headers) + " |",
        "|" + "|".join("---" for _ in headers) + "|",
    ]
    for case in cases:
        lines.append("| " + " | ".join(_markdown_cell(value) for value in case.table_row(include_metadata=include_metadata)) + " |")
    return "\n".join(lines) + "\n"


def export_cases(path: Path, cases: Sequence[TestCase], format_name: str, *, include_metadata: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if format_name in {"table", "markdown"}:
        path.write_text(render_table(cases, include_metadata=include_metadata), encoding="utf-8")
        return
    if format_name == "json":
        path.write_text(json.dumps([asdict(case) for case in cases], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return
    if format_name == "csv":
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(CANONICAL_HEADERS + (METADATA_HEADERS if include_metadata else []))
            writer.writerows(case.table_row(include_metadata=include_metadata) for case in cases)
        return
    if format_name == "xlsx":
        _export_xlsx(path, cases, include_metadata=include_metadata)
        return
    raise ValueError(f"неподдерживаемый формат: {format_name}")


def _export_xlsx(path: Path, cases: Sequence[TestCase], *, include_metadata: bool = False) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Тест-кейсы"
    headers = CANONICAL_HEADERS + (METADATA_HEADERS if include_metadata else [])
    worksheet.append(headers)
    for case in cases:
        worksheet.append(case.table_row(include_metadata=include_metadata))
    # All exported fields are literal text, including values starting with '='.
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.data_type = "s"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10, color="222222")
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    widths = (42, 54, 54, 54, 54) + ((24, 32, 16, 24) if include_metadata else ())
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False
    if cases:
        table = Table(displayName="TestCasesTable", ref=f"A1:{get_column_letter(len(headers))}{worksheet.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        worksheet.add_table(table)
    workbook.save(path)
