# ADE: исследование рынка и конкурентов

Дата: 14 октября 2026. Метод: WebSearch/WebFetch. Каждое утверждение идёт со ссылкой; номера в квадратных скобках ведут к списку источников в конце.
Пометка **«гипотеза»** означает вывод, который не подтверждён источником напрямую. Если цена или функция не найдены, так и написано: «н/д».
Оговорка: часть источников — блоги конкурентов (Deckary, Prezi, Presentations.AI, ChatSlide, Plus AI). Их оценки помечены как вендорские.

---

## 1. Краткий вывод (TL;DR)

1. **Главная угроза — «агенты внутри PowerPoint».** За 2026 год их выпустили Microsoft, Anthropic и OpenAI:
   - Copilot Agent Mode правит существующий дек «сохраняя форматирование, структуру и брендинг» (GA февраль 2026) [3], а с июня 2026 умеет взять стиль приложенного дека [4];
   - Claude for PowerPoint строит слайды «из ваших клиентских/корпоративных шаблонов» по slide master [20];
   - ChatGPT for PowerPoint (бета, май 2026) правит существующие слайды [21].

   Все трое работают с **layouts/placeholders шаблона**. Ни один источник не подтверждает надёжную **структурную адаптацию готового дизайнерского слайда** (добавить 4-ю карточку в композицию из 3, растянуть процесс с 4 до 5 шагов при сохранении сетки), а также пакетный визуальный QA.
2. **Copilot прямо признаёт ограничения:**
   - он «не меняет размер шрифта и плейсхолдера» под длинный текст [2][6];
   - опирается на sample slides и плейсхолдеры, а шаблоны на текстовых боксах не понимает [6][7];
   - официальный список известных проблем: наложения текста, незаполненные плейсхолдеры, слишком мелкий текст [2];
   - представитель Microsoft на форуме: «Copilot just isn't built for that kind of precise formatting work yet» [9].
3. **Сервисы-генераторы (Gamma, Beautiful.ai, Decktopus, Pitch, Canva, Figma Slides) работают внутри своей системы тем.** Импорт PPTX даёт им только цвета, шрифты и лого, но не макеты [12][13].
   - У Gamma фича-реквест «генерировать по лейаутам моего шаблона» набрал 113 голосов [14].
   - Экспорт в PPTX у этих сервисов теряет slide master, градиенты и шрифты [11][28][32].
4. **Корпоративные платформы шаблонов (Templafy, UpSlide, Macabacus, Prezent) продают «бренд-комплаенс» крупным компаниям.** Они либо сами делают шаблоны (Prezent: от $399 за пользователя в месяц [19]), либо требуют внедрения. Задача «макет + новый контент → адаптированный дизайн» у них решена только частично (Prezent Template Converter перекрашивает дек в бренд [36]).
5. **Генераторы документов по шаблону (Plumsail, Nintex, S-Docs, Aspose) умеют только размножать строки таблиц по тегам.** Вложенные коллекции не поддерживаются, таблица не переносится на следующий слайд, нужна ручная разметка и код [24][25][26]. Карточки, процессы и смысловую адаптацию они не решают.
6. **Боль подтверждена количественно:**
   - ~7 ч в неделю на PowerPoint, 37% этого времени уходит на форматирование, 47% презентаций не соответствуют бренду (n=1102, 2020) [27];
   - аналитики консалтинга тратят 40–60% времени на слайды (малая выборка n≈20, 2026) [29];
   - сервисы-исполнители берут $13–66 за слайд [30], фрилансеры — $25–150 в час [31].
7. **Ниша ADE** (гипотеза): «детерминированная адаптация чужого готового дизайна под новый контент с визуальным QA». Это слой между агентами в PowerPoint (гибко, но неточно) и шаблонизаторами документов (точно, но негибко). Рынок ПО для презентаций — $8,3–13,4 млрд в 2025 году, рост 12–17% в год [33][34]. Gamma при ARR >$100 млн [35] показывает, что за AI-презентации платят.

---

## 2. Сравнительная таблица

Легенда:
- **Чужой шаблон** — умеет ли инструмент работать с существующим PPTX-шаблоном или макетом клиента, а не со своими темами.
- **Структурная адаптация** — умеет ли добавить карточку, строку таблицы, категорию диаграммы или шаг процесса, сохранив дизайн.
- **Редактируемый PPTX** — получается ли на выходе нативный редактируемый PowerPoint.

| Продукт | Чужой шаблон | Структурная адаптация | Редактируемый PPTX | Цена | ЦА |
|---|---|---|---|---|---|
| **MS Copilot в PowerPoint** | Да: корпоративный шаблон или brand kit; ≥4 лейаута, ≤40 лейаутов, .potx [2][6]; агент правит существующий дек [3]; стиль приложенного дека (GA июнь 2026) [4] | Частично: выбирает лейауты; не меняет размеры шрифтов и плейсхолдеров [6]; игнорирует указанный лейаут [8]; не повторяет точное форматирование образца [9] | Да, нативно | Add-on $18/польз./мес (Business), $30 (Enterprise) + базовая лицензия M365 [37] | Корпорации на M365 |
| **Gemini в Google Slides** | Частично: «приложить дек как референс стиля»; полный импорт темы не описан [5] | н/д | Нативные Google Slides; PPTX только через экспорт Slides (не проверено) | Входит в Workspace Business/Enterprise и Google AI Pro/Ultra [5] | Пользователи Workspace; на старте только en-US [5] |
| **Claude for PowerPoint** (Anthropic) | Да: читает slide master, лейауты, шрифты и цвета; «build new slides using your existing client or corporate templates» [20] | Частично: правки отдельных слайдов, нативные диаграммы, реструктуризация [20]; у сложных шаблонов в обзоре потребовались мелкие ручные правки, графики «тесные» [22] | Да, нативно | Входит в Claude Pro/Max/Team/Enterprise [20] | Консалтинг, корпорации |
| **ChatGPT for PowerPoint** (OpenAI) | Правит существующие деки [21]; работа с шаблонами не описана | н/д | Да, нативно (add-in) | Бета, доступна всем аккаунтам, включая бесплатные [21] | Широкая |
| **Gamma** | Слабо: импорт .pptx как **темы** (цвета, шрифты, лого) [12]; лейауты не переносятся: «borrow parts of it but not use the actual layout» [13] | Нет для чужого макета; карточки гибкие только внутри своей системы | Частично: таблицы редактируемые, slide master пустой, градиенты в один цвет [11]; карточки выше 16:9 вылезают за слайд [38] | Plus $8–10, Pro $15–20, Ultra $39–100/мес [39] | Массовый рынок; ~70 млн пользователей [35] |
| **Beautiful.ai** | Частично: свой конструктор тем, импорт PowerPoint [15]; «custom brand requirements… hit guardrails quickly» (вендорский обзор Prezi) [40] | Да, но только внутри своих Smart Slides (авто-раскладка) | Да, но шрифты слетают (3 из 15 слайдов в тесте), анимации теряются [41] | Pro $14,50/мес; Team $40/польз./мес [15] | SMB, маркетинг, продажи |
| **Plus AI** | Свой шаблон только в Enterprise: «build slides from your organization's existing slides and slide masters» [16][42] | Remix: «pick a slide template and format your content» [42]; правит деки «не ломая шаблон» (вендорский обзор Prezi) [40] | Да, работает внутри PowerPoint и Google Slides | $10–30/мес; Max $200; Enterprise по запросу [16] | Консультанты, корпорации |
| **SlidesAI** | н/д; темы «в основном ещё недоступны» (вендорский обзор) [43] | н/д | Экспорт PPTX, на выходе бывают расхождения [28] | $0 / $10 / $20 в мес [43] | Образование, SMB |
| **Tome** | Продукт презентаций **закрыт 30.04.2025**, компания ушла в CRM (Lightfield) [44] | — | — | — | — |
| **Pitch** | н/д (свои шаблоны) | н/д | н/д | Plus $13–15; Team $19–23/место; Business $25–30 [45] | Стартапы, команды |
| **Presentations.AI** | Brand Sync берёт шрифты, цвета и лого **с сайта**; импорт файла-шаблона не описан [46] | н/д | «editable .pptx» (заявление вендора) [46] | Бесплатный старт, платные тарифы (цены н/д) [46] | Стартапы, продажи, консультанты |
| **Decktopus** | Нет: бренд только как цвета, шрифты, лого (Business) [47] | Нет | «formatting losses with complex layouts» (вендорский обзор) [47] | Pro $14,99/мес; Business $34,99/польз./мес [47] | SMB |
| **Prezent.ai** | Шаблоны бренда делает сам вендор («expert-designed brand templates»); Template Converter переводит любой дек в бренд [19][36] | Частично: перекраска и приведение к шаблону; адаптация структуры н/д | Редактируемость результата на странице не указана [36] | Lite $399/польз./мес; остальные тарифы по запросу [19] | Enterprise (коммуникации, фарма, финансы) |
| **Templafy** | Да: шаблоны, лейауты, бренд-правила клиента подключаются к агенту [17][18] | Агент в PowerPoint: новые слайды, правка форматирования, редактируемые графики [18]; глубина н/д | Да, внутри PowerPoint | Агент бесплатен для индивидуальных пользователей; корпоративная часть по запросу [18] | Крупные корпорации (KPMG, BDO), >4 млн пользователей [18] |
| **UpSlide** | Да: шаблоны и библиотека фирмы; позиционируется как слой «после ChatGPT/Claude»: доводит AI-черновики до бренда [48] | Форматирование в один клик, связка Excel→PPT; генерации нет | Да | н/д | M&A, банки, PE, консалтинг [48] |
| **think-cell** | Работает внутри шаблона клиента (надстройка) | Графики: 40+ типов, связка с Excel; AI-функций нет [49] | Да | ~$230–327/польз./год (оценка) [49] | Консалтинг, банки |
| **Macabacus** | Работает внутри шаблона клиента (надстройка) | Форматирование, библиотеки слайдов, бренд-проверка; AI нет [50] | Да | От $200/польз./год [50] | Инвестбанки, консалтинг |
| **Storydoc** | Нет (интерактивные веб-деки) | Нет | Экспорт в PDF; PPTX не подтверждён [51] | $19 / $36 в мес [51] | Продажи, маркетинг |
| **Canva** | Импорт PPTX ломает slide master, SmartArt, шрифты [52]; Brand Kit в Magic Design только для цветов (вендорский обзор Plus AI) [53]; Autofill API заполняет **тегированные поля** brand template [54] | Нет: Autofill описывает только заполнение существующих полей [54] | Экспорт PPTX «может сплющивать сложные дизайны» [28]; в тесте «corrupted the design» (вендорский обзор) [53] | Autofill от Canva Pro/Teams/Enterprise [54] | Маркетинг, SMB |
| **Figma Slides** | AI-аутлайн с кастомным шаблоном: не использует слайды шаблона, меняет кегли, результат каждый раз разный (форум, апрель 2026, без ответа) [55] | Нет | Экспорт PPTX: шрифты заменяются, градиенты в один цвет, интерактив становится картинками [32] | н/д | Дизайнеры, продуктовые команды |
| **Manus** | Да: загрузка PowerPoint-шаблона (ноябрь 2025) [56] | н/д | н/д | н/д | Широкая |
| **Skywork** | Да: «upload your own PPTX or PDF template» [57] | н/д | Формат выгрузки не указан [57] | н/д | Широкая |
| **SlideSpeak** | Да: API с PPTX/POTX-шаблонами клиента, Brand MCP для Claude [58] | н/д | Да (PPTX) | н/д | Разработчики, интеграции |
| **2Slides (Fast PPT)** | Да: заполнение слотов шаблона; «does not improvise a new layout»; онбординг шаблона ~48 ч [59] | **Нет**: новых лейаутов не создаёт [59] | Да (PPTX) | 10 кредитов за страницу [59] | SMB |
| **Felo Slides** | Да: «upload your own template and let AI fill the content» [60] | н/д; плохо работает с шаблонами на свободных текст-боксах [60] | н/д | н/д | Широкая |
| **Unique AI (PPTX skill)** | Да: дублирует слайды реального шаблона из базы знаний и **делает визуальный QA** (влезание текста, наложения, контраст) [23] | Ограничено: «template is the ceiling»; чего нет в шаблоне, того не будет [23] | Да (.pptx) | Корпоративная платформа (н/д) | Финансы, enterprise |
| **Plumsail / Nintex / S-Docs** | Да: свой шаблон с тегами [24][25][26] | Только повтор строк и таблиц; вложенные коллекции нельзя [25]; таблица не переносится на следующий слайд [26] | Да | Plumsail: триал 30 дней (цены н/д) [24] | Автоматизация отчётов (Power Automate, Salesforce) |
| **Aspose.Slides** | Да, через код (SDK) | Всё, что напишет разработчик; шейпы адресуются по индексу [61] | Да | Лицензия SDK (н/д) | Разработчики |
| **Сервисы-исполнители (24Slides, SlideGenius)** | Да: принимают мастер и брендбук клиента [30] | Да, вручную | Да | $13–66 за слайд [30]; $25–150 в час, дек 15–20 слайдов за $800–3000 [31] | Корпорации, маркетинг |

---

## 3. Свидетельства реальной потребности

### 3.1 Жалобы на Copilot и корпоративные шаблоны
- **Февраль–март 2024.** Copilot игнорирует брендовые шаблоны и отвечает «Sorry I can't do that». Обходной путь — открыть пустой шаблон и генерировать в нём. Asset Library Copilot не читает [10].
- **Январь 2025.** Copilot «almost always reuses the same layout», игнорирует явно названный лейаут. Microsoft: требования (лого всегда слева сверху, заголовок сверху) «may be beyond Copilot's current capabilities» [8].
- **Июнь 2025.** Пользователю нужно перестроить дек точно по образцу: шрифты, позиции, новые слайды при переполнении. Ответ: «Copilot just isn't built for that kind of precise formatting work yet» [9].
- **Июль 2025.** В десктопе нельзя выбрать кастомный шаблон при создании через Copilot; Microsoft называет это «known limitation» [1].
- **Официальная документация Microsoft:**
  - известные проблемы — «overlapping text, images over content, color problems, unfilled placeholders, text that is too small» [2];
  - Copilot «does not change the font size or placeholder size» [6];
  - текст плейсхолдера Copilot воспринимает только как стиль [6][7].
- **Вывод:** Microsoft перекладывает работу на клиента. Чтобы Copilot справлялся, шаблон надо перестроить под Copilot: sample slides, правильные плейсхолдеры [6][7]. Отсюда возникает рынок консалтинга вида «Copilot-ready template» (SlideGenius и др.) [7]. Гипотеза: это же создаёт спрос на инструмент, которому перестройка шаблона не нужна.

### 3.2 Жалобы на Gamma
- Реквест «Allow me to generate from a template or the layouts in my template»: **113 голосов**. Цитаты пользователей:
  - «I MUST comply with our branding requirements»;
  - «For Gamma to be ready for business, we need the ability to create templates like PowerPoint»;
  - Gamma использует загруженный шаблон «as inspiration rather than applying them strictly» [14].
- Реквест «Add ability to upload branding guidelines or PPT templates»: 38 голосов. Gamma выпустила импорт **тем**, но пользователи пишут: «What we need is… our own templates — not themes» [63].
- «Improve theme import…»: Gamma «borrow[s] parts of it but not use the actual layout I've provided»; «still doing its own thing» [13].
- Обзор консультанта (август 2025): в PPTX-экспорте нет slide master; в шаблонах нет «title + subtitle» под ключевые сообщения [38].

### 3.3 Другие AI-инструменты
- Figma Slides AI с кастомным шаблоном: каждый раз разный результат, шаблонные слайды не используются, кегли меняются (апрель 2026, 0 ответов) [55].
- Canva Magic Design: Brand Kit применяет только палитру; экспорт PPTX «corrupted the design» (вендорский обзор Plus AI) [53].
- Beautiful.ai: в тесте шрифты слетели на 3 из 15 слайдов; сетка мешает точно поставить лого [41].

### 3.4 Время и деньги
- **Empower, 2020 (n=1102, Германия и США):**
  - ~7 ч в неделю на создание презентаций;
  - **37%** этого времени (~2,6 ч в неделю) уходит на форматирование;
  - **47%** презентаций не соответствуют корпоративному дизайну [27].
- **SlideFlow (Wharton, июнь 2026, опрос ~20 аналитиков):**
  - 40–60% рабочей недели уходит на слайды;
  - 79% оценили полезность AI на 4–5 из 5;
  - цитата: «modification to content requires the whole slide to be taken apart and rebuilt»;
  - вывод авторов: ни один конкурент не совмещает интеграцию с PowerPoint, контроль slide master и консалтинговые сценарии [29].
- **Вторичная сводка** (не проверено по первоисточникам): 37% времени на форматирование, 42% на поиск шаблонов и старых слайдов, 38% на обновление и ребрендинг деков; верификация AI-результата съедает часть экономии [62].
- **Дообработка после AI:** у Copilot до клиентской готовности «typical 15–20 minutes of manual polish» (вендорская оценка Prezi) [40]; у Felo 5–10 минут [60].
- **Templafy:** экономия >30% времени на повторяющемся контенте (заявление вендора) [18].
- **UpSlide:** «without hours of reformatting»; клиент сократил подготовку питча с 2 ч до ≤30 мин (отзыв) [48].
- **Цена ручного труда:**
  - 24Slides: Fix Up $13–16, Redesign $35–43, Redraw $53–66 за слайд [30];
  - SlideGenius: дизайнер $25–150 в час, агентство от $150 в час, дек 15–20 слайдов за $800–3000 [31].

### 3.5 Размер рынка
- **Presentation software:**
  - TBRC: $7,04 млрд (2024) → $8,29 млрд (2025) → $15,76 млрд (2029), CAGR 17,4% [33];
  - SkyQuest: $12,0 млрд (2024) → $13,43 млрд (2025) → $33,01 млрд (2033), CAGR 11,9% [34];
  - оценки расходятся почти вдвое, поэтому для бизнес-плана брать с осторожностью.
- **Gamma:** ARR >$100 млн, ~70 млн пользователей, оценка $2,1 млрд (ноябрь 2025) [35].
- **Рынок агентств и сервисов дизайна презентаций:** надёжной цифры не найдено. Гипотеза: оценивать снизу вверх, например число консультантов и банкиров × часы на форматирование × ставка.

---

## 4. Что дизайнеры презентаций до сих пор делают вручную

Источники: [9][23][29][30][31][6][59].
1. **Перенос контента из Word или черновика в фирменный макет.** Это базовая услуга агентств (Fix Up и Redesign у 24Slides [30]). Copilot и Claude делают это лишь частично: не подгоняют размеры и не держат точное форматирование [6][9].
2. **Подгонка композиции под другое количество элементов:** 3→4 карточки, 4→5 шагов, новая строка или колонка таблицы. Шаблонизаторы умеют только повтор строк [25][26]. «Fast PPT» прямо не создаёт новых лейаутов [59]. Аналитики говорят, что слайд приходится «разбирать и собирать заново» [29].
3. **Длинный текст:** перенос, сокращение, уменьшение кегля в пределах гайдлайна, разбиение на 2 слайда. Copilot размеры не меняет [6]; S-Docs советует резать данные вручную через LIMIT/OFFSET [26].
4. **Обновление диаграмм:** новые категории и цифры. Частично закрыто think-cell и UpSlide через связку с Excel [48][49]; без Excel остаётся ручная работа (гипотеза).
5. **Визуальная проверка:** наложения, выходы за поля, мелкий текст, контраст. Есть только у отдельных корпоративных агентов, например Unique AI [23]; Copilot называет это известными проблемами [2].
6. **Приведение «AI-черновика» к бренду.** Этим целиком позиционируется UpSlide [48]; Prezent продаёт Template Converter [36].
7. **Построение «Copilot-ready» шаблона:** sample slides, плейсхолдеры вместо текст-боксов. Это новая ручная работа, которую создал сам Copilot [6][7].

---

## 5. Выводы для ADE

**Позиционирование.** «Отдай готовый дизайн и новый Word — получи тот же дизайн с новым контентом, где структура выросла или сжалась, со штампом визуального QA». Отличия от конкурентов:
- **Работа с дизайн-композицией, а не только с плейсхолдерами.** Copilot и 2Slides требуют правильных плейсхолдеров и не создают новых лейаутов [6][59]. Профессиональные макеты часто собраны из групп и шейпов (гипотеза; косвенно подтверждается советом Microsoft перестроить текст-боксы в плейсхолдеры [7]).
- **Структурная адаптация с сохранением сетки** (N→N+1 карточек, шагов, строк). Ни у одного конкурента не подтверждена [1–61].
- **Детерминизм и пакетность.** У Figma AI результат каждый раз разный [55]. Гипотеза: агентства и отделы отчётности ценят воспроизводимость.
- **Визуальный QA как часть продукта, а не рекомендация «проверьте сами»** [2][20].

**Угрозы:**
- Claude, Copilot и ChatGPT в PowerPoint быстро улучшаются (релизы 2026 года [3][4][20][21]).
- Окно возможностей — **гипотеза: 12–24 мес.** до того, как агенты научатся структурной адаптации «из коробки».
- Защита: качество QA на сложных макетах, бенчмарк точности, интеграция как skill или MCP для самих этих агентов. SlideSpeak уже делает Brand MCP для Claude [58]. UpSlide позиционирует себя как слой «после ChatGPT/Claude» [48].

**Сегменты для проверки (гипотезы для интервью):**
1. Студии и фрилансеры дизайна презентаций. Продают Fix Up и Redesign за $13–43 за слайд [30]; ADE снижает их себестоимость.
2. Консалтинг и банкинг: аналитики, 40–60% времени на слайды [29]. Платят за think-cell и Macabacus ~$200–330 в год [49][50].
3. Корпоративные коммуникации и отчётность: регулярные отчёты по фиксированному макету с меняющимся объёмом. Сегодня это Plumsail, Nintex, S-Docs с их ограничениями [24][25][26].

**Ценовой ориентир (гипотеза).** Ниже стоимости ручной работы ($13+ за слайд [30]) и сопоставимо с надстройками ($10–40 в месяц [15][16]). Enterprise-ориентир: Prezent от $399 в месяц [19].

**Что проверить дальше:**
- Сделать реальные тесты на одном и том же макете (3→4 карточки, 4→5 шагов): Copilot Agent Mode, Claude for PowerPoint, ChatGPT for PowerPoint, Gemini, Plus AI. Публичных бенчмарков структурной адаптации не найдено.
- Найти количественные данные по рынку сервисов дизайна презентаций.

---

## 6. Источники

1. https://learn.microsoft.com/en-us/answers/questions/5436522/not-able-to-select-custom-template-when-trying-to
2. https://support.microsoft.com/office/use-your-organization-s-branding-with-copilot-in-powerpoint-c8bc6df5-37ed-4398-8b90-f78a8fdcf9bb
3. https://mc.merill.net/message/RM548520
4. https://mc.merill.net/message/RM555880
5. https://workspaceupdates.googleblog.com/2026/06/create-fully-native-and-editable-presentations-with-Gemini-in-Google-Slides.html
6. https://support.microsoft.com/topic/046c23d5-012e-49e0-8579-fe49302959fc
7. https://www.slidegenius.com/blog/copilot/powerpoint-placeholders-vs-text-boxes-copilot
8. https://learn.microsoft.com/en-us/answers/a/1908366
9. https://techcommunity.microsoft.com/discussions/microsoft-copilot/is-copilot-capable-of-rebuilding-an-existing-powerpoint-presentation-based-on-sa/4419479
10. https://techcommunity.microsoft.com/discussions/microsoft365copilot/powerpoint-company-branding/4072079
11. https://help.gamma.app/en/articles/8022861-how-do-i-export-my-gamma
12. https://help.gamma.app/en/articles/11029150-can-i-create-a-custom-theme-for-my-gamma-workspace
13. https://ideas.gamma.app/ideas/p/improve-theme-import-brand-import-template-import-or-quality
14. https://ideas.gamma.app/ideas/p/corporate-template-support
15. https://www.beautiful.ai/pricing
16. https://plusai.com/pricing
17. https://www.templafy.com/ai-presentation-maker
18. https://macaubusiness.com/templafy-launches-free-document-agent-in-microsoft-marketplace/
19. https://prezent.ai/pricing
20. https://claude.com/docs/office-agents/powerpoint.md
21. https://gigazine.net/gsc_news/en/20260522-chatgpt-for-powerpoint
22. https://slidemodel.com/claude-for-powerpoint/
23. https://docs.unique.ai/administrators/skills-in-unique-ai-conduct/pptx-skill-presentations-from-company-templates
24. https://plumsail.com/blog/create-powerpoint-presentations-from-pptx-templates/
25. https://help.nintex.com/en-US/nwc/Content/Designer/Actions/GenerateDocument-RepeatDataPowerPoint.htm
26. https://help.sdocs.com/sdocs/template-architecture/document-formats/pptx-format/pptx-rendering-considerations
27. https://www.empowersuite.com/hubfs/Marketing/Downloads/PowerPoint%20Studie%202020/Englisch/The%20Ultimate%20Global%20PowerPoint%20Study%20-%20empower.pdf
28. https://www.chatslide.ai/articles/which-ai-slide-generator-creates-editable-ppt-files
29. https://fisher.wharton.upenn.edu/wp-content/uploads/2026/06/SlideFlow-writeup.pdf
30. https://24slides.com/pricing
31. https://www.slidegenius.com/cm-faq-question/what-is-the-typical-hourly-rate-for-a-presentation-designer
32. https://figma-signup.helpjuice.com/import-and-export-2/export-from-figma-slides
33. https://www.gii.tw/report/tbrc1706136-presentation-software-global-market-report.html
34. https://www.gii.tw/report/sky1973346-presentation-software-market-size-share-growth.html
35. https://pulse2.com/gamma-68-million-at-2-1-billion-valuation-raised-as-it-surpasses-100-million-in-arr/amp/
36. https://www.prezent.ai/template-converter
37. https://eesel.ai/blog/copilot-pricing
38. https://colleagueinterrupted.substack.com/p/can-gamma-ai-replace-powerpoint-i
39. https://costbench.com/software/ai-presentations/gamma
40. https://prezi.com/blog/client-decks-ai-slide-builder-consultants-2026/
41. https://dupple.com/reviews/beautiful-ai
42. https://plusai.com/use-cases/ai-for-consultants
43. https://deckary.com/blog/slidesai-review
44. https://www.chatslide.ai/articles/tome-ai-shut-down-what-happened
45. https://www.costbench.com/software/ai-presentations/pitch/
46. https://www.presentations.ai/ai-instructions
47. https://deckary.com/blog/decktopus-review
48. https://www.upslide.com/
49. https://deckary.com/blog/think-cell-pricing
50. https://deckary.com/blog/macabacus-review
51. https://www.capterra.com/p/10036236/Storydoc/
52. https://slidemodel.com/edit-powerpoint-templates-in-canva/
53. https://plusai.com/blog/canva-magic-design-for-presentations-and-alternative-ai-presentation-tools
54. https://canva.dev/docs/connect/autofill-guide
55. https://forum.figma.com/ask-the-community-7/using-create-figma-slides-outline-ai-with-a-custom-template-52518
56. https://alternativeto.net/news/2025/11/create-on-brand-slides-with-manus-new-powerpoint-template-feature/
57. https://skywork.ai/help/tutorial/slides
58. https://docs.slidespeak.co/basics/custom-templates
59. https://2slides.com/blog/on-brand-ai-slides-template-fill
60. https://felo.ai/blog/ai-to-create-slides-from-template/
61. https://tutorials.aspose.com/slides/net/presentation-manipulation/perform-mail-merge-in-presentations/
62. https://intuitionlabs.ai/articles/ai-agents-powerpoint-problem-2026
63. https://ideas.gamma.app/ideas/p/corporate-branding-theme-import

Дополнительно просмотрено: https://support.microsoft.com/copilot-powerpoint (создание из Word: до 5 файлов, Word Styles, до 24 МБ), https://artificialanalysis.ai/zh/agents/presentations (каталог из 16 AI-инструментов для презентаций, без оценок), https://www.parse.gl/prompts/p/i-want-a-presentation-generator-that-uses-my-corporate-brand-template-and-does-not-generate-random-designs--4a13b53c-55fa-4434-b7b7-eccb43423b9d (на запрос «генератор по моему корпоративному шаблону» AI-поисковики чаще всего советуют Beautiful.ai, Gamma и Canva, то есть инструменты, которые как раз плохо держат чужой шаблон; это наблюдение — гипотеза о незакрытой нише).

Ограничения исследования: Reddit через поисковый прокси оказался недоступен (запросы с фильтром домена отклонены). Голос пользователей собран с форумов Microsoft, Gamma Ideas и Figma Forum.
