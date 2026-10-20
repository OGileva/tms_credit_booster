# Наборы

Правила `include` и `exclude` поддерживают `case_ids`, `products`, `sections`, `tags`, `statuses`. Внутри блока правила объединяются по ИЛИ; `exclude` применяется после `include`.

Примеры:

```bash
qa-cases export --suite smoke
qa-cases export --suite full-regression --format xlsx
```
