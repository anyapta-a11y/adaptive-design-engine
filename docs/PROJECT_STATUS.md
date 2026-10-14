# Статус проекта (читать первым, ≤60 строк)

Обновлено: 2026-10-14.

## Фаза
1 — прототип v0 по концепции D-005 (библиотека слайдов + skill + сборка). Анна дала направление 2026-10-14.
Прототип в `ade/`; полный продукт — после демонстрации прототипа.

## Ключевые выводы
- Адаптации (таблица, диаграмма, карточки, процесс, длинный текст) на чистом макете реализуемы: `experiments/spike_01`.
- Риск «понять чужой макет» снят концепцией D-005: библиотека своя и размечена. Спрос: `docs/research/demand.md`; конкуренты: `docs/research/competitors-deep.md` — прямого аналога не найдено.
- Главная рыночная угроза — Copilot Agent Mode / Claude for PowerPoint / ChatGPT for PowerPoint (2026).
- Graphify и Obsidian Skills не внедряем (D-001, D-004).

## Где что лежит
- Решения: `docs/DECISIONS.md` · Задачи: `docs/BACKLOG.md`
- Исследования: `docs/research/{market,platforms,token-economy}.md`
- Эксперименты: `experiments/spike_01/` (RESULTS.md, QA_REPORT.md)
- Агенты: `.claude/agents/`

## Где остановились (2026-10-14)
- Прототип v0.2 в `ade/`: каталог, сборка, QA; вертикальные списки, свободные элементы, нарисованные столбцы. 36 тестов.
- Сквозной пример: `examples/northern/` (Word → 8 слайдов, QA PASS). Skill: `.claude/skills/ade-build/`.
- Библиотека Анны (экспорт Canva, 10 слайдов) лежит ТОЛЬКО у неё: `Агент Шаблон\test\`. В git не класть без её согласия.
- В работе у неё на компьютере: копия `test\library_native.pptx` — слайды 3, 5, 8 перерисовываются нативными диаграммой/таблицами (решение Анны). Статус неизвестен: сессия на её компьютере оборвалась.
- Следующее: разметить её библиотеку (LIBRARY_GUIDE), прогнать Word-документ локально у неё.

## Как запустить локально
`pip install python-pptx pillow lxml python-docx` → `python examples/make_library.py examples/library.pptx` → `python -m ade build --library examples/library.pptx --plan examples/plan_demo.json --out out.pptx` → `python -m ade qa out.pptx --plan examples/plan_demo.json`.

## Открытые вопросы к владельцу
1. GO на ворота A–C или другое решение?
2. Доступ к PowerPoint (desktop/web) и Copilot/Claude for PowerPoint для Gate B.
3. Контакты для 5 интервью (Gate C).

## Метрики эффективности (вести по задачам)
| Дата | Задача | Агент | Токены | Вызовы | Итерации QA |
|---|---|---|---|---|---|
| 2026-10-14 | Рынок | researcher | 164k | 125 | — |
| 2026-10-14 | Платформы | researcher | 169k | 130 | — |
| 2026-10-14 | Graphify/Obsidian | researcher | 101k | 16 | — |
| 2026-10-14 | spike_01 исправления | developer | 27k | 8 | 2 |
| 2026-10-14 | spike_01 QA ×2 | qa | 80k + 18k | 19 | — |
