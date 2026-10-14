# Статус проекта (читать первым, ≤60 строк)

Обновлено: 2026-10-14.

## Фаза
0 — исследование целесообразности завершено. **Ждём решения Анны** по `docs/FEASIBILITY_STUDY.md` (рекомендация: REVISE — сузить до PPTX, пройти ворота A–C до MVP).
Продуктовый код не писать до решения.

## Ключевые выводы
- Адаптации (таблица, диаграмма, карточки, процесс, длинный текст) на чистом макете реализуемы: `experiments/spike_01`.
- Главный нерешённый риск — распознавание структуры чужого макета (не проверено).
- Главная рыночная угроза — Copilot Agent Mode / Claude for PowerPoint / ChatGPT for PowerPoint (2026).
- Graphify и Obsidian Skills не внедряем (D-001, D-004).

## Где что лежит
- Решения: `docs/DECISIONS.md` · Задачи: `docs/BACKLOG.md`
- Исследования: `docs/research/{market,platforms,token-economy}.md`
- Эксперименты: `experiments/spike_01/` (RESULTS.md, QA_REPORT.md)
- Агенты: `.claude/agents/`

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
