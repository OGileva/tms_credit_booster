from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Sequence

from .core import (
    create_section,
    export_cases,
    import_cases,
    load_sections,
    load_suite,
    render_table,
    resolve_section,
    resolve_tags,
    select_cases,
    validate_repository,
    write_index,
)


FORMATS = ("table", "xlsx", "csv", "json", "markdown")


def _root(value: str) -> Path:
    return Path(value).expanduser().resolve()


def _add_filters(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--suite", help="имя набора из suites/")
    parser.add_argument("--product", action="append", default=[], help="продукт или glob; можно повторять")
    parser.add_argument("--section", action="append", default=[], help="путь или название функционала; можно повторять")
    parser.add_argument("--tag", action="append", default=[], help="обязательный тег; можно повторять")
    parser.add_argument("--id", action="append", default=[], help="точный ID; можно повторять")
    parser.add_argument("--search", help="подстрока во всём содержимом кейса")
    parser.add_argument("--status", choices=("active", "draft", "deprecated", "all"), default="active")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="qa-cases", description="Текстовый Git-каталог тест-кейсов")
    parser.add_argument("--root", type=_root, default=Path.cwd(), help="корень репозитория")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("validate", help="проверить структуру и содержимое всех кейсов")
    commands.add_parser("index", help="пересобрать generated/index.json")
    commands.add_parser("sections", help="показать доступные продукты и разделы")

    list_parser = commands.add_parser("list", help="показать краткий список кейсов")
    _add_filters(list_parser)

    for name, help_text in (
        ("export", "выгрузить выборку; по умолчанию Markdown-таблица"),
        ("select", "псевдоним команды export"),
    ):
        export_parser = commands.add_parser(name, help=help_text)
        _add_filters(export_parser)
        export_parser.add_argument("--format", choices=FORMATS, default="table")
        export_parser.add_argument("--output", type=Path, help="путь файла; table без output печатается в терминал")

    import_parser = commands.add_parser("import", help="массово добавить кейсы из приложенного файла")
    import_parser.add_argument("--source", required=True, type=Path)
    import_parser.add_argument("--section", required=True, help="путь, название или alias раздела")
    import_parser.add_argument("--product", help="продукт для устранения неоднозначности")
    import_parser.add_argument("--tag", action="append", default=[], help="тег для всех импортируемых кейсов")
    import_parser.add_argument("--task", action="append", default=[], help="ID требования или задачи")
    import_parser.add_argument("--status", choices=("active", "draft"), default="active")
    import_parser.add_argument("--dry-run", action="store_true", help="проверить без записи файлов")

    section_parser = commands.add_parser("create-section", help="создать новый раздел функционала")
    section_parser.add_argument("--product", required=True)
    section_parser.add_argument("--path", required=True)
    section_parser.add_argument("--name", required=True)
    section_parser.add_argument("--id-prefix", required=True)
    section_parser.add_argument("--alias", action="append", default=[])
    return parser


def _checked_cases(root: Path):
    cases, issues = validate_repository(root)
    for issue in issues:
        print(issue.render(), file=sys.stderr)
    errors = [issue for issue in issues if issue.level == "error"]
    if errors:
        raise ValueError(f"валидация не пройдена: ошибок {len(errors)}")
    return cases


def _filtered(root: Path, args: argparse.Namespace):
    cases = _checked_cases(root)
    suite = load_suite(root, args.suite) if args.suite else None
    tags = resolve_tags(root, args.tag) if args.tag else ()
    return select_cases(
        cases,
        suite=suite,
        products=args.product,
        sections=args.section,
        tags=tags,
        ids=args.id,
        search=args.search,
        status=args.status,
    )


def _default_output(root: Path, format_name: str) -> Path:
    extension = "md" if format_name in {"table", "markdown"} else format_name
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return root / "exports" / f"selection-{timestamp}.{extension}"


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    root = args.root.resolve()
    if not (root / "project.json").exists():
        parser.error(f"не найден project.json: {root}")
    try:
        if args.command == "validate":
            cases = _checked_cases(root)
            print(f"OK: кейсов {len(cases)}, разделов {len(load_sections(root))}")
            return 0
        if args.command == "index":
            cases = _checked_cases(root)
            output = write_index(root, cases)
            print(f"Индекс обновлён: {output.relative_to(root)}; кейсов {len(cases)}")
            return 0
        if args.command == "sections":
            print("Product\tSection\tName\tAliases\tID prefix")
            for section in load_sections(root):
                print(f"{section.product}\t{section.path}\t{section.name}\t{', '.join(section.aliases)}\t{section.id_prefix}")
            return 0
        if args.command == "list":
            cases = _filtered(root, args)
            print("ID\tProduct\tSection\tTags\tName\tFile")
            for case in cases:
                print(f"{case.id}\t{case.product}\t{case.section_path}\t{','.join(case.tags)}\t{case.title}\t{case.file}")
            print(f"Найдено: {len(cases)}", file=sys.stderr)
            return 0
        if args.command in {"export", "select"}:
            cases = _filtered(root, args)
            if not cases:
                raise ValueError("по заданным условиям кейсы не найдены")
            if args.format == "table" and args.output is None:
                print(render_table(cases), end="")
                return 0
            output = args.output or _default_output(root, args.format)
            if not output.is_absolute():
                output = root / output
            export_cases(output, cases, args.format)
            print(f"Выгружено: {len(cases)}; файл: {output}")
            return 0
        if args.command == "import":
            source = args.source.expanduser().resolve()
            if not source.exists():
                raise ValueError(f"файл не найден: {source}")
            section = resolve_section(root, args.section, args.product)
            cases = import_cases(
                root,
                source_path=source,
                section=section,
                command_tags=args.tag,
                status=args.status,
                tasks=args.task,
                dry_run=args.dry_run,
            )
            mode = "Проверено без записи" if args.dry_run else "Добавлено"
            print(f"{mode}: {len(cases)}; раздел: {section.product}/{section.path}")
            for case in cases:
                print(f"{case.id}\t{case.title}\t{','.join(case.tags)}")
            return 0
        if args.command == "create-section":
            section = create_section(
                root,
                product=args.product,
                path=args.path,
                name=args.name,
                id_prefix=args.id_prefix,
                aliases=args.alias,
            )
            print(f"Создан раздел: {section.product}/{section.path}; prefix={section.id_prefix}")
            return 0
    except (OSError, ValueError, KeyError, json.JSONDecodeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
