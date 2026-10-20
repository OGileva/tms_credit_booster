# QA test cases

- **Status:** draft
- **Date:** 2026-10-20
- **Owner:** QATL

Git-репозиторий ручных тест-кейсов с массовым вводом через Excel и воспроизводимыми выборками для регресса или отдельного функционала.

---

## Быстрый старт

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -e .

qa-cases new-pack --product credit-booster --function payments --name CBN-1234
qa-cases validate
qa-cases index
```

Новый файл появится в `cases/credit-booster/payments/CBN-1234.xlsx`. Заполните строки на листе `Тест-кейсы`, затем повторите `validate` и `index`.

---

## Как хранить кейсы

```text
cases/
└── credit-booster/
    ├── payments/
    │   ├── CBN-1234.xlsx
    │   └── CBN-1300.xlsx
    └── scoring/
        └── CBN-1400.xlsx
```

Путь задаёт продукт и функционал. Внутри XLSX остаются только пять колонок из `AGENTS.md`; служебные теги в файл не добавляются.

Рекомендуемый `Name`:

```text
CB-PAY-0001 — Успешное проведение платежа
```

ID стабилен и используется в suites. Название после тире можно уточнять.

---

## Команды

| Команда | Результат |
|---|---|
| `qa-cases validate` | Проверяет все исходные XLSX и дубли ID/названий |
| `qa-cases index` | Пересобирает `generated/index.json` |
| `qa-cases list --function payments` | Показывает кейсы функционала |
| `qa-cases list --search таймаут` | Ищет текст во всех пяти полях |
| `qa-cases select --function payments --output exports/payments.xlsx` | Выгружает функциональную выборку |
| `qa-cases select --suite full-regression --output exports/regression.xlsx` | Выгружает suite |
| `qa-cases new-pack ...` | Создаёт новый пакет из Excel-шаблона |

Фильтры `--product`, `--function`, `--id` и `--search` можно комбинировать. Несколько значений одного фильтра задаются повторением параметра.

---

## Наборы регресса

- `full-regression.json` включает все кейсы.
- `smoke.json` и `release-critical.json` изначально пусты: их нужно наполнить подтверждёнными ID или правилами по функционалам.
- Правила suite поддерживают `case_ids`, `globs`, `products`, `functionalities` и блок исключений.

Пример:

```json
{
  "id": "payments-regression",
  "description": "Регресс платежей",
  "include": {
    "case_ids": ["CB-PAY-0001"],
    "globs": ["cases/credit-booster/payments/*.xlsx"],
    "products": [],
    "functionalities": []
  },
  "exclude": {
    "case_ids": [],
    "globs": [],
    "products": [],
    "functionalities": []
  }
}
```

---

## Совместная работа

1. Создать отдельную ветку.
2. Добавить или изменить исходный XLSX.
3. Обновить suite, если состав набора изменился.
4. Выполнить `make check test index`.
5. В Merge Request указать задачу, функционал, добавленные/изменённые ID и влияние на регресс.

Подробно: [`regulations/reglament-test-cases-workflow.md`](regulations/reglament-test-cases-workflow.md).

---

## Разделы

| Раздел | Назначение |
|---|---|
| [`cases/`](cases/) | Исходные тест-кейсы |
| [`suites/`](suites/) | Наборы регресса |
| [`_templates/`](_templates/) | Excel-шаблон |
| [`regulations/`](regulations/) | Регламент изменений |
