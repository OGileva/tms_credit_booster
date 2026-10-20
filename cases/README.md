# Каталог тест-кейсов

Продукт — папка первого уровня. Вложенный путь описывает функционал. Конечный раздел содержит `section.json` и файлы `<ID>.md`.

Новый раздел:

```bash
qa-cases create-section \
  --product credit-booster \
  --path scoring/loan-history \
  --name "История займов" \
  --id-prefix CB-SCR-LH \
  --alias "История кредитов"
```
