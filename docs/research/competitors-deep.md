# Конкуренты: углублённая проверка «делает ли кто-то то же, что ADE»

Дата: 14.10.2026. Дополняет `market.md` (не повторяет). Метод: WebSearch + WebFetch страниц вендоров и обзоров.
Вопрос: берёт ли продукт **чужой готовый PPTX-макет** + новый контент, **меняет число элементов** (3→4 карточки, 5→7 строк, 3→4 категории, 4→5 этапов, длинный текст) и отдаёт **нативный редактируемый PPTX**?
Метки: точно / вероятно / гипотеза. «Н/д» = не найдено в открытых источниках, это не равно «не умеет».

## 0. Ограничения проверки (честно)
- YouTube-демо и транскрипты через WebFetch/WebSearch **не получены**. Видео-доказательств в этом отчёте нет. Для Copilot/Claude/ChatGPT нужен собственный тест (Gate B).
- Платных аккаунтов нет, ни один продукт не запускался. Всё ниже: документация, release notes, обзоры. Прямых скриншотов «3→4 карточки» на чужом макете **не найдено ни у одного продукта**.
- Reddit и соцсети недоступны.

## 1. Итоговая таблица

Столбцы: (а) чужой макет клиента; (б) меняет число элементов в макете; (в) нативный редактируемый PPTX; (г) уровень доказательства.

| Продукт | (а) | (б) | (в) | (г) |
|---|---|---|---|---|
| Copilot Agent Mode (PowerPoint) | Да, brand kit/шаблон организации | Н/д; только общее «сохраняя форматирование, структуру, брендинг» | Да | Документация + заявление; демо структурной адаптации не найдено |
| Claude for PowerPoint | Да, читает slide master | Н/д; примеры: «restructure the storyline», «bullets → process flow» | Да, нативные диаграммы | Документация; один обзор с тестом: правки вручную, «jammed» графики |
| ChatGPT for PowerPoint | Не описано | Н/д | Да (add-in) | Слабое: сторонние блоги, без тестов |
| Gemini in Slides | Только «attach deck as style reference» | Не описано | Нативные Google Slides; PPTX-экспорт в посте не упомянут | Документация (пост Google) |
| Plus AI (Remix) | Да (Teams/Enterprise, «template conversion») | **Вероятно да, но не доказано**: Reformat переносит слайд в другой layout, текст переписывается | Да (внутри PPT/Slides) | Документация, без примеров изменения числа элементов |
| Beautiful.ai | Импорт PPTX -> **Classic Slides** (не Smart) | Smart Slides: да, но только свои | Экспорт: фигуры вместо текстбоксов | Документация + обзоры |
| think-cell | Работает в шаблоне клиента | **Нет** (Assist: текст, перевод, чарты из данных) | Да | Документация/пресс-релиз |
| Prezent (Template Converter) | Любой дек -> шаблон бренда | Н/д (перекраска/приведение к бренду) | Н/д | Маркетинг |
| Templafy | Да (шаблоны клиента) | Н/д | Да | Маркетинг, демо недоступно |
| UpSlide | Да | Не AI-структурирование (оформление в 1 клик) | Да | Маркетинг (market.md) |
| SlideSpeak (API) | Да, .pptx/.potx | Н/д | Да | Документация, без деталей |
| 2Slides | Да | **Нет** (по заявлению вендора: «does not improvise a new layout») | Да | Вендорский блог (market.md) |
| Canva Autofill | Brand template (в Canva) | **Нет**: только text/image поля | Экспорт PPTX с потерями | Документация API |
| Gamma | Импорт: **только текст**, без стилей и layout | Нет для чужого макета | Частично | Официальный help |
| Presenton (open-source) | Да: слайды PPTX -> layouts | Не объяснено; рекомендуют лимиты символов | Не подтверждено | Документация |
| PPTAgent (open-source, EMNLP 2025) | Да: референс-дек, кластеризация layout | **Да, механика есть**: clone_paragraph / del_span | PPTX (формулировка в статье неоднозначна) | Научная статья + репо |
| Kimi / Manus / Genspark | Manus: шаблон PPTX (market.md); Kimi, Genspark: н/д | Н/д | Kimi: заявляют editable PPTX (блог) | Н/д или маркетинг |
| Presentations.AI | Шрифты/цвета с сайта, не макеты | Н/д | Заявление вендора | Маркетинг |
| Pitch | Н/д | Н/д | Н/д | Нет данных |

## 2. Детали по кандидатам

### 2.1 Beautiful.ai (особый запрос)
- Механика Smart Slides **есть**: «Add a fourth bullet point to a three-item list and the other three resize» [B3]; «Drop in a third bullet point and the spacing redistributes» [B2]. Точно (обзоры; официальный маркетинг подтверждает идею «design automation» [B4]).
- Но **с чужим PPTX она не работает**: официальный блог: импортированные PowerPoint-слайды появляются как **Classic Slides** («you can now import them into Beautiful.ai as Classic Slides»), то есть свободный холст без авто-раскладки; Smart Slide можно превратить в Classic в один клик, обратный путь в источнике не описан [B4]. Help-статья по импорту не говорит, что слайды становятся Smart [B1]; блог обещает «fully editable and inherit from your presentation's theme wherever possible» [B5]. Вывод: **вероятно** импорт = нередактируемая авто-раскладкой «картинка из фигур», перераспределения нет. Дата блога про Classic Slides не установлена, поведение сегодня могло измениться (проверка триалом).
- Slide AI умеет «Switch structures with a single click» и использовать текущий контент слайда как основу, но про Classic/импортированные слайды в статье ничего нет [B6]. Н/д.
- Экспорт в PPTX: «Elements are converted to shape objects rather than native PowerPoint text boxes» [B3]; 3 из 15 слайдов с проблемами шрифтов, анимации потеряны [B2]. Сам вендор публикует «how to export to ppt and why you shouldn't» [B7].
- Итог: механика-аналог ADE внутри закрытой экосистемы; **чужой макет + нативный PPTX на выходе = нет** (вероятно).

### 2.2 Агенты в Office (Copilot, Claude, ChatGPT, Gemini)
- **Copilot Agent Mode.** Release notes MS: на существующем деке может обновлять контент, улучшать layout «preserving formatting, structure, and branding», GA март 2026, desktop [C1]. Support-страница: «creating a slide using your template's layouts», перекраска под шаблон [C2]. Структурной адаптации (+1 карточка) нигде не описано. Ограничения из market.md (не меняет размеры плейсхолдеров/шрифта [6]) остаются единственным прямым свидетельством: **не опровергнуто и не подтверждено для Agent Mode** (market.md ссылается на более старый Copilot). Вероятно, потребуется свой тест.
- **Claude for PowerPoint.** Официальная документация: «Build new slides using your existing client or corporate templates», «Restructure the storyline across slides 4 to 7», «Turn these bullets into a process flow diagram», нативные диаграммы [CL1]. Это **создание новых слайдов по layouts**, а не перестройка дизайнерской композиции чужого слайда. Ограничения самих документов: «Not recommended for final client deliverables without human review» [CL1]. Независимый тест (SlideModel): на шаблоне нужны «minor fixes - font sizing, color, or arranging shapes», при генерации с нуля «jammed» графики и неверные таблицы; тест, в котором добавляли бы/убирали элементы, **не проводился** (в промпте это было запрещено) [CL2]. Точно.
- **ChatGPT for PowerPoint.** Единственный найденный разбор (блог MindStudio, вендор конкурирующих автоматизаций, без тестов) утверждает: бренд-цвета/шрифты не применяются автоматически, правки идут по выбранному слайду, таблицы/диаграммы из Word «come through as text» [CG1]. Доказательство **слабое**. Сравнительная статья того же автора без тестов [CG2].
- **Gemini in Slides.** Пост Google (30.06.2026): генерация нативных редактируемых презентаций, можно приложить дек как референс стиля; редактирование существующих слайдов, смена layout и экспорт в PPTX в посте **не упомянуты**; en-US only [G1]. Точно по содержимому поста.

### 2.3 Plus AI
- Remix: Reformat «moves a slide into a different layout», Synthesize, Elaborate; режимы Default (текст переписывается, сокращается или расширяется) и Strict (текст сохраняется, но может не влезть) [P1][P2]. Работает «into any template»; Teams импортируют свои шаблоны, Enterprise включает «white-glove template conversion» [P1].
- **Не описано** прямо: добавление/удаление карточек, шагов, строк. Страница Remix на этот вопрос не отвечает [P2]. Вероятно, Reformat ближе всех из коммерческих продуктов к цели, но Strict-режим сам признаёт, что вёрстка может не влезть. Гипотеза: «шаблон» у Plus = набор layouts, подготовленных вендором (конверсия на стороне Plus), а не любой дизайнерский макет.
- Демо/видео: не получены.

### 2.4 Генераторы «по шаблону»
- **Gamma:** официальный help: «Right now we only support importing the text, not the styling or layouts» [GM1]. Точно: чужой макет не используется.
- **Canva Autofill:** поля только `text` и `image` в примере, видео не поддерживается, внешние URL не поддерживаются, про число строк/элементов в документации ничего [CA1]. Структурной адаптации нет (точно по документации).
- **SlideSpeak:** API принимает .pptx/.potx; как обрабатываются счётчики элементов, не сказано [S1]. **2Slides:** заполнение слотов, нового layout не создаёт (вендор, market.md [59]).
- **Presenton (open-source):** в Template Studio загружается PPTX, каждый слайд становится layout, анализируются текстовые и графические поля [PR1]. Как layout переживает разное число элементов — не описано; рекомендуют лимиты символов и тесты на низкой/высокой плотности [PR1]. Это шаблонный подход «слоты с лимитами», а не перестройка композиции (вероятно).
- **Kimi Slides, Genspark, Manus:** страницы help не говорят о загрузке собственного PPTX-шаблона (Kimi [K1]); у Genspark в выдаче только сторонние сравнения [GS1], не вскрыты. Н/д. Manus: загрузка PPTX-шаблона заявлена (market.md [56]) без деталей.

### 2.5 PPTAgent (ближайший по механике)
- Статья (arXiv 2501.03936): этап I — кластеризация слайдов референса по layout (текст заменяется плейсхолдерами, картинки — цветом) и извлечение схемы элементов; этап II — LLM выдаёт вызовы API `replace_span`, `replace_image`, `clone_paragraph`, `del_span`, `del_image`, которые **правят слайд референса**. Если нужно больше абзацев, чем в образце, агент клонирует абзац; если меньше — удаляет с конца [A1].
- Это **прямой аналог идеи** «N → N+1 элемент через клонирование», но: (1) операции на уровне абзацев/спанов, **без перераспределения геометрии** (в статье этого нет; успех >95% — это «выполнился код», не «выглядит хорошо»); (2) авторы сами пишут, что оверлеи возможны, дизайн «not fully optimized using visual information»; (3) таблицы/диаграммы в операциях не упомянуты [A1].
- Репозиторий (MIT, ~5k звёзд, DeepPresenter принят на ACL 2026): вывод PPTX, Docker, нет Windows, рекомендуется собственная 9B-модель [A2]. Уровень: научная работа + код, **не продукт**; на чужих сложных макетах результат не проверен нами.

### 2.6 Корпоративные надстройки
- **think-cell:** вопреки `market.md` («AI-функций нет») с **1 июня 2026** есть **think-cell Assist**: чат, данные Morningstar/Statista -> чарты, перевод «keeping layout», резюме и action titles, AI-правка текста (rewrite/shorten/restructure text) [T1][T2]. Нет описания перестройки раскладки или смены числа элементов; «user keeps control over wording, structure, and formatting» [T1]. Точно: перестройки композиции нет. Исправить market.md.
- **Prezent Template Converter:** «analyzing fonts, colors, images, and footers, then updating every slide to match your brand»; на вход любой дек, на выход «entirely new deck using the template you've chosen»; про число элементов и редактируемость страница молчит [PZ1].
- **Templafy:** «Automatically adapt to user input… brand elements and business context»; про смену раскладки по числу элементов страница не говорит; ссылка на демо не раскрыта [TF1]. Маркетинг.
- **UpSlide, Unique AI:** см. `market.md` (Unique: «template is the ceiling»). Новых данных не найдено (release notes март 2026 про PPTX не пишут [U1]).

## 3. Вывод

**Прямого аналога не найдено.** Ни у одного из ~20 кандидатов нет открытого доказательства (демо, help, обзор со скриншотами), что продукт на **чужом дизайнерском PPTX** меняет 3→4 карточки / 5→7 строк / 4→5 этапов с перераспределением геометрии, сохраняет стиль и отдаёт нативный PPTX. Уверенность: **вероятно** (отсутствие доказательств ≠ доказательство отсутствия; видео и закрытые функции не проверены).

**Кто ближе всех (по убыванию):**
1. **Агенты Office (Copilot Agent Mode, Claude for PowerPoint)** — тот же сценарий и тот же формат, но документированы операции «по layouts шаблона», а не перестройка композиции; независимый тест Claude показал ручную доводку. Реальная угроза, но **не проверена на нашем бенчмарке** -> Gate B.
2. **Plus AI Remix** — Reformat в другой layout внутри PPT, кастомные шаблоны. Не описано изменение числа элементов; режим Strict сам оговаривает риск «не влезет».
3. **PPTAgent** (open-source) — единственная найденная механика «клонировать/удалить элементы образца». Не продукт, без геометрии и QA.
4. **Beautiful.ai** — наилучшая механика авто-перераспределения, но только в Smart Slides; импорт PPTX даёт Classic Slides (вероятно, без авто-раскладки), экспорт — фигуры вместо текстбоксов.

**Чем отличается ADE (если ниша существует):** (1) вход — произвольный готовый макет из групп/фигур, без перестройки под плейсхолдеры; (2) структурная перестройка с пересчётом сетки (N->N+1, таблица, диаграмма) вместо слотов с лимитами; (3) нативный PPTX; (4) автоматический структурный и визуальный QA, которого ни один продукт не декларирует как часть выдачи (кроме Unique AI на уровне визуального QA).

**Ниша:** вероятно есть, но узкая и недолговечная. Главные риски: Copilot/Claude могут уже справляться (не проверено), PPTAgent-подобные идеи открыты, главный технический риск самого ADE — распознавание структуры чужого макета.

**Следующие шаги (проверяемые):**
1. Gate B: один и тот же синтетический макет, сценарии 3→4, 4→5, 5→7, 3→4 категории и длинный текст; прогнать Copilot Agent Mode, Claude for PowerPoint, ChatGPT for PowerPoint, Plus AI (триал). Сохранить исходные и итоговые PPTX.
2. Триал Beautiful.ai: импортировать синтетический PPTX, добавить элемент, посмотреть, остаётся ли Classic.
3. Запустить PPTAgent в пустой временной директории на нашем макете (только после решения владельца по Фазе 0/экспериментам).
4. Исправить в `market.md`: think-cell Assist (июнь 2026).

## 4. Источники
- B1 https://support.beautiful.ai/hc/en-us/articles/360059654532 · B2 https://dupple.com/reviews/beautiful-ai · B3 https://www.fast.io/resources/beautiful-ai-review-2026.md · B4 https://beautiful.ai/blog/introducing-classic-slides · B5 https://www.beautiful.ai/blog/4-easy-ways-to-cut-ties-with-powerpoint · B6 https://support.beautiful.ai/hc/en-us/articles/43350069148557 · B7 https://www.beautiful.ai/blog/how-to-export-to-ppt-and-why-you-shouldnt (заголовок из выдачи, страницу не открывал)
- C1 https://m365admin.handsontek.net/?p=19490 · C2 https://support.microsoft.com/topic/agent-mode-in-powerpoint-frontier-008f17aa-8e5f-4cda-ba6c-0588000bdad7 (страница не упоминает «Agent Mode» в тексте, описывает Copilot in PowerPoint) · статья про Agent Mode (пример с Excel-диаграммами): https://pasqualepillitteri.it/en/news/1401/microsoft-copilot-agent-mode-word-excel-powerpoint-april-2026
- CL1 https://claude.com/docs/office-agents/powerpoint.md · CL2 https://slidemodel.com/claude-for-powerpoint/
- CG1 https://www.mindstudio.ai/blog/chatgpt-in-powerpoint-build-editable-decks · CG2 https://www.mindstudio.ai/blog/chatgpt-powerpoint-vs-copilot-vs-claude-comparison
- G1 https://workspaceupdates.googleblog.com/2026/06/create-fully-native-and-editable-presentations-with-Gemini-in-Google-Slides.html
- P1 https://plusai.com/features/reformat-slides-with-remix · P2 https://guide.plusai.com/ai-for-presentations/remix-slides
- GM1 https://help.gamma.app/en/articles/7856738-importing-slides-and-documents · CA1 https://canva.dev/docs/connect/autofill-guide · S1 https://docs.slidespeak.co/basics/custom-templates · PR1 https://docs.presenton.ai/user-guide/branding-and-design/templates · K1 https://www.kimi.com/help/slides/ppt-overview · GS1 https://plusai.com/blog/in-depth-genspark-review (в выдаче, не открывал)
- A1 https://arxiv.org/pdf/2501.03936 · A2 https://github.com/icip-cas/PPTAgent
- T1 https://www.thinkcell.com/en/presentation-ai · T2 https://www.schneider.im/think-cell-assist-generative-ai-inside-the-powerpoint-and-excel-workflow/ (сторонний пересказ, 03.06.2026)
- PZ1 https://www.prezent.ai/template-converter · TF1 https://www.templafy.com/ai-presentation-maker · U1 https://unique.ai/en/blog/unique-ai-release-notes-march-2026
