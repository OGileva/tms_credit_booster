# Наборы тест-кейсов

JSON-файлы содержат правила включения и исключения. Все поля внутри `include` и `exclude` объединяются по логике ИЛИ; блок `exclude` применяется после включения.

Поддерживаются:

- `case_ids` — точные стабильные ID;
- `globs` — шаблоны пути к исходным XLSX;
- `products` — шаблоны продукта;
- `functionalities` — шаблоны функционала.

Проверка набора:

```bash
qa-cases select --suite smoke --output exports/smoke.xlsx
```
