from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from .core import (
    create_pack,
    export_cases,
    load_config,
    load_suite,
    select_cases,
    validate_repository,
    write_index,
)


def _root(value: str) -> Path:
    return Path(value).expanduser().resolve()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="qa-cases", description="Управление Excel-каталогом тест-кейсов")
    parser.add_argument("--root", type=_root, default=Path.cwd(), help="корень репозитория")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("validate", help="проверить все исходные XLSX")
    subparsers.add_parser("index", help="пересобрать generated/index.json")

    list_parser = subparsers.add_parser("list", help="показать выбранные кейсы")
    _add_filters(list_parser)

    select_parser = subparsers.add_parser("select", help="выгрузить выбранные кейсы в XLSX")
    _add_filters(select_parser)
    select_parser.add_argument("--output", required=True, type=Path, help="путь итогового XLSX")

    new_parser = subparsers.add_parser("new-pack", help="создать пакет из шаблона")
    new_parser.add_argument("--product", required=True)
    new_parser.add_argument("--function", dest="functionality", required=True)
    new_parser.add_argument("--name", required=True)
    return parser


def _add_filters(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--suite", help="имя JSON-файла из suites/")
    parser.add_argument("--product", action="append", default=[], help="product или glob; можно повторять")
    parser.add_argument("--function", dest="functionalities", action="append", default=[], help="function или glob; можно повторять")
    parser.add_argument("--id", action="append", default=[], help="точный ID; можно повторять")
    parser.add_argument("--search", help="подстрока во всех полях")


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
    return select_cases(
        cases,
        suite=suite,
        products=args.product,
        functionalities=args.functionalities,
        ids=args.id,
        search=args.search,
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    root = args.root.resolve()
    if not (root / "project.json").exists():
        parser.error(f"не найден project.json в корне: {root}")

    try:
        if args.command == "validate":
            cases = _checked_cases(root)
            print(f"OK: файлов {len(set(case.source for case in cases))}, кейсов {len(cases)}")
            return 0
        if args.command == "index":
            cases = _checked_cases(root)
            output = write_index(root, cases)
            print(f"Индекс обновлён: {output.relative_to(root)}; кейсов {len(cases)}")
            return 0
        if args.command == "list":
            cases = _filtered(root, args)
            print("ID\tProduct\tFunction\tName\tSource")
            for case in cases:
                print(f"{case.id}\t{case.product}\t{case.functionality}\t{case.name}\t{case.source}:{case.row}")
            print(f"Найдено: {len(cases)}", file=sys.stderr)
            return 0
        if args.command == "select":
            cases = _filtered(root, args)
            if not cases:
                raise ValueError("по заданным условиям не найдено ни одного тест-кейса")
            output = args.output
            if not output.is_absolute():
                output = root / output
            export_cases(output, cases, load_config(root))
            print(f"Выгружено: {len(cases)}; файл: {output}")
            return 0
        if args.command == "new-pack":
            output = create_pack(root, args.product, args.functionality, args.name)
            print(f"Создан пакет: {output.relative_to(root)}")
            return 0
    except (OSError, ValueError, KeyError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
