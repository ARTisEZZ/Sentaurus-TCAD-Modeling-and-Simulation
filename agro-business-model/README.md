# Финмодель сельхозбизнеса (Томская область)

Расчёт бизнес-модели: купить участок, построить дом и рядом хозяйство
(кормовые поля, коровы, прочие животные, огород) — с оценкой инвестиций
и окупаемости по трём схемам финансирования.

## Состав

| Файл | Назначение |
|---|---|
| `model.py` | Финмодель: CAPEX, OPEX, выручка, юнит-экономика, кэшфлоу, NPV/IRR/окупаемость, безубыточность, чувствительность. Чистый Python, без зависимостей. |
| `params.example.json` | Шаблон параметров. Скопируйте в `params.json` и правьте под свои данные. |
| `.mcp.json` | Конфигурация MCP-серверов для сбора данных и работы с таблицами. |
| `report.md`, `cashflow.csv` | Генерируются при запуске `model.py`. |

Скиллы лежат в `../.claude/skills/`: `agro-capex`, `agro-unit-economics`,
`agro-opex-revenue`, `agro-investment`.

## Быстрый старт

```bash
cd agro-business-model
python3 model.py          # печатает отчёт, пишет report.md и cashflow.csv
```

Чтобы подставить свои цифры:

```bash
cp params.example.json params.json   # затем отредактируйте params.json
python3 model.py
```

## Установка MCP-серверов (в интерактивном Claude Code)

В этой облачной сессии команды `claude mcp add` не выполняются — запускайте
их у себя. Минимальный набор — `fetch` + `excel`; остальное по желанию.

```bash
# Сбор внешних данных (цены земли, стройки, кормов, условия грантов)
claude mcp add fetch -- uvx mcp-server-fetch
claude mcp add brave -e BRAVE_API_KEY=ВАШ_КЛЮЧ -- npx -y @modelcontextprotocol/server-brave-search

# Таблицы и хранение сценариев
claude mcp add excel -- uvx excel-mcp-server
claude mcp add fs -- npx -y @modelcontextprotocol/server-filesystem ./agro-business-model
claude mcp add db -- uvx mcp-server-sqlite --db-path ./agro-business-model/agro.db

# Структурирование многошагового расчёта (опционально)
claude mcp add think -- npx -y @modelcontextprotocol/server-sequential-thinking
```

Либо положите `.mcp.json` (из этой папки) в корень проекта — Claude Code
подхватит серверы автоматически (не забудьте вписать `BRAVE_API_KEY`).

Проверка: `claude mcp list`.

## Рабочий процесс

1. Через `fetch`/`brave-search` собрать актуальные вводные по Томску
   (цена участка ИЖС и сельхозземли, ₽/м² стройки, цена молока, корма,
   условия гранта «Агростартап» и льготного кредита Россельхозбанка).
2. Внести их в `params.json`.
3. Прогнать скиллы по порядку: `agro-capex` → `agro-unit-economics` →
   `agro-opex-revenue` → `agro-investment` (или просто `python3 model.py`).
4. Сравнить сценарии финансирования, выгрузить в Excel через MCP `excel`.

## Что показывает базовый расчёт

- Сравнение **трёх схем**: свои средства / грант КФХ + свои / льготный кредит.
- Грант даёт лучшую окупаемость, кредит сдвигает положительный поток на
  более поздние годы, «свои средства» — без долга, но самый долгий возврат.
- **Важно**: дом — личный актив; он раздувает CAPEX и ухудшает NPV «бизнеса».
  Для чистой бизнес-оценки вынесите дом из CAPEX (обнулите `house_*` и
  `utilities`) и считайте жильё отдельно.

> Все цифры в модели **иллюстративные** (оценка 2026 г.). Перед решением
> подтвердите их актуальными данными по вашему участку и рынку.
