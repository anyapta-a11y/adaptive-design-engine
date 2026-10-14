# ADE: технические возможности платформ (PPTX, Google Slides, Figma)

Дата исследования: 2026-10-14. Источники: прежде всего официальная документация. Всё, что не подтверждено источником, помечено **гипотеза**.

Операции ADE, которые проверялись:
1. замена текста;
2. добавление строк таблицы с сохранением стиля;
3. добавление карточки (дублирование группы фигур и перераспределение);
4. изменение числа категорий нативной диаграммы;
5. автоподбор размеров текстовых блоков;
6. расширение процессной схемы с 4 до 5 этапов (SmartArt или фигуры + коннекторы).

Сквозное требование: результат остаётся редактируемым, объекты не подменяются картинками.

---

## Краткий вывод

| Платформа | Чем сильна | Главные ограничения для ADE |
|---|---|---|
| PPTX (OOXML) | Полная объектная модель: нативные диаграммы со встроенным xlsx, таблицы, группы, SmartArt, layouts и masters. Офлайн-обработка файла | Нет движка вёрстки. `normAutofit fontScale` вычисляет PowerPoint, а не библиотека. python-pptx не умеет добавлять строки таблиц, клонировать фигуры и править SmartArt. Рендер LibreOffice отличается от PowerPoint |
| Google Slides API | Облачный batchUpdate: `insertTableRows`, `duplicateObject`, `groupObjects`, `replaceAllText` | Нет нативных диаграмм: только диаграммы из Google Sheets. Значения autofit (`fontScale`) только для чтения. При импорте PPTX диаграммы становятся картинками. Квота 60 записей в минуту на пользователя |
| Figma и Figma Slides | Auto-layout и компоненты отлично подходят для «добавить карточку». Plugin API и MCP `use_figma` дают полную запись | REST API в основном только читает. Нативных диаграмм нет. Таблицы есть только в Slides/FigJam. Figma Slides не поддерживает компоненты, стили и переменные. Импорт PPTX теряет таблицы, диаграммы и шрифты |

Рекомендация для MVP: **PPTX как основной формат**, прямые правки OOXML (lxml поверх python-pptx). Для рендер-проверки нужен рендерер, близкий к PowerPoint. Google Slides и Figma подключать позже как отдельные адаптеры IDM, с явной матрицей деградаций.

---

## A. PowerPoint / PPTX (OOXML)

### A.1 python-pptx

- **Таблицы.** `add_table` создаёт новую таблицу ([API shapes](https://python-pptx.readthedocs.io/en/latest/api/shapes.html)). Метода добавления строки нет: это давний feature request [#86](https://github.com/scanny/python-pptx/issues/86), предложенный [PR #399](https://github.com/scanny/python-pptx/pull/399) статус не подтвердил. Обходной путь: deepcopy элемента `<a:tr>` через lxml, затем поправить высоту `graphicFrame` (**гипотеза**: стандартная практика сообщества, официального подтверждения нет).
- **Диаграммы.** `chart.replace_data(chart_data)` переписывает и XML диаграммы, и встроенный Excel-лист ([API chart](https://python-pptx.readthedocs.io/en/latest/api/chart.html)), поэтому число категорий меняется.
  - При увеличении числа серий новые серии получают цвет последней ([#539](https://github.com/scanny/python-pptx/issues/539), не решено). Обход: явные цвета или «запасные» серии в шаблоне.
  - `replace_data` теряет числовой формат, если не передать `number_format` ([#666](https://github.com/scanny/python-pptx/issues/666)).
  - 3D-диаграммы не поддерживаются: «Most chart types other than 3D types are supported» ([user/charts](https://python-pptx.readthedocs.io/en/latest/user/charts.html)).
- **Группы.** `add_group_shape` есть в `SlideShapes` и в `GroupShapes`. Метода клонирования фигуры нет ([API shapes](https://python-pptx.readthedocs.io/en/latest/api/shapes.html)). Дублирование карточки делается через deepcopy XML и новые `id` у `cNvPr`. Для картинок нужно копировать или переиспользовать связи (rId) (**гипотеза**: инженерная практика).
- **SmartArt.** Не поддерживается. Для SmartArt `shape_type` возвращает `None`, специального свойства нет ([graphfrm source](https://python-pptx.readthedocs.io/en/latest/_modules/pptx/shapes/graphfrm.html)). Feature request [#83](https://github.com/scanny/python-pptx/issues/83) открыт. Фактически SmartArt доступен только как сырой XML.
- **Autofit.**
  - `fit_text()` сам подбирает один размер шрифта по метрикам TTF-файла и выставляет `noAutofit` ([analysis: fit text](https://python-pptx.readthedocs.io/en/latest/dev/analysis/txt-fit-text.html)).
  - Ограничения `fit_text()`: путь к шрифту указывается вручную, шрифт и начертание предполагаются одинаковыми по всему тексту, разрывы строк схлопываются.
  - `normAutofit` с `fontScale` python-pptx **не вычисляет**: значение пересчитывает PowerPoint при открытии или редактировании (**гипотеза**, косвенно подтверждена тем, что спецификация хранит `fontScale` как готовое значение; см. A.5).

### A.2 Office.js PowerPoint API (состояние на 2025–2026)

Требования по наборам API ([requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets?view=word-js-1.7)):
- 1.4: фигуры (Windows 2207).
- 1.8: bindings, shapes, tables (Windows 2504).
- 1.9: форматирование таблиц (Windows 2508).
- 1.10: accessibility, фон, гиперссылки (Windows 2601).
- Windows LTSC поддерживает только 1.1–1.5, iPad только 1.1.

| Возможность | Статус | Источник |
|---|---|---|
| `addTable`, `addGroup` | 1.8 | [ShapeCollection](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection) |
| `rows.add(index,count)`, `deleteRows` | 1.9 | [TableRowCollection](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablerowcollection) |
| `mergeCells`, `clear`, `styleSettings` | 1.9 | [Table](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.table) |
| `addChart` | **нет** | [ShapeCollection](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection) |
| Тип `Chart`/`SmartArt` | только распознаётся | [ShapeType](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapetype) |
| `textFrame.autoSizeSetting` | 1.4: AutoSizeNone / TextToFitShape / ShapeToFitText | [TextFrame](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.textframe) |
| `insertSlidesFromBase64` | KeepSourceFormatting / UseDestinationTheme | [Insert slides](https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/insert-slides-into-presentation) |
| `addPicture` | preview | [Preview APIs](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-preview-apis?view=outlook-js-1.7) |

Вывод: Office.js подходит как надстройка, работающая в живом PowerPoint. Главный плюс: PowerPoint сам считает autofit и рендер. Но править данные диаграмм и SmartArt через Office.js нельзя. Гибридный путь: сервер собирает PPTX, затем `insertSlidesFromBase64` вставляет слайды в открытый документ.

### A.3 Aspose.Slides (коммерческая библиотека)

- **Таблицы.** `rows.add_clone(row, False)` и `insert_clone(index, row, False)` клонируют строку вместе с форматированием ([rows & columns](https://docs.aspose.com/slides/python-net/manage-rows-and-columns/)). Это готовая операция «добавить строку в стиле».
- **SmartArt.** Узлы редактируются: `AllNodes.AddNode()`, `ChildNodes.AddNodeByPosition()`, `RemoveNode`, `node.TextFrame.Text` ([SmartArt nodes](https://docs.aspose.com/slides/net/manage-smartart-shape-node/)). Тип layout задаётся только при создании и потом доступен лишь для чтения ([SmartArt](https://docs.aspose.com/slides/python-net/manage-smartart-shape/)). Отсюда путь для 4 → 5 этапов: SmartArt.AddNode.
- **Диаграммы.** `ChartDataWorkbook` даёт чтение и запись ячеек и внешние книги. После замены workbook нужно вручную очистить старые series и categories. `.xlsb` не поддерживается ([chart workbook](https://docs.aspose.com/slides/python-net/chart-workbook/)).
- **Autofit.** Документация описывает поведение `TextAutofitType.Normal`, но не говорит, вычисляет ли библиотека `fontScale` сама. Документация предупреждает, что подмена шрифта меняет метрики ([autofit](https://docs.aspose.com/slides/net/manage-autofit-settings/index.md)). Что Aspose пересчитывает fontScale при рендере: **гипотеза**, нужно проверить тестом.
- Есть собственный рендер в PNG и PDF с настройкой шрифтов ([custom fonts](https://docs.aspose.com/slides/python-net/custom-font/)). Точность относительно PowerPoint: **гипотеза**, нужен бенчмарк.

### A.4 LibreOffice как рендерер

- **SmartArt.** Выпущенные версии импортируют не сам SmartArt, а запасной рисунок `ppt/diagrams/drawing1.xml`. Собственный layout-движок LibreOffice автор называет «unacceptable». Прогресс работ по SmartArt оценивается примерно в 33–66% ([Collabora forum](https://forum.collaboraonline.com/t/problems-with-ooxml-import-export-of-smartart/4702)). Последствие: если ADE меняет `data.xml`, но не перегенерирует `drawing.xml`, LibreOffice покажет старую картинку.
- **Шрифты и autofit.** В трекере есть многочисленные баги о расхождении размеров текста и autofit, например [bug 132895](https://bugs.libreoffice.org/show_bug.cgi?id=132895) и [bug 149997](https://www.libreoffice.org/bugzilla/show_bug.cgi?id=149997). Количественной оценки точности в официальных источниках нет. Вывод «LO-рендер пригоден для грубой проверки переполнения, но не как эталон пиксельной точности»: **гипотеза**.
- Без установленных фирменных шрифтов любой рендерер (LO, Aspose, Google) подставляет другие шрифты, и меняются переносы (см. FAQ Aspose выше). ADE должен везти шрифты с собой.

### A.5 Особенности OOXML, важные для ADE

- **Диаграммы.** XML диаграммы хранит кэш значений (`c:numCache`/`c:strCache`), а связанный встроенный xlsx служит источником для «Edit Data». Обновлять нужно оба, иначе PowerPoint при редактировании перезапишет кэш из xlsx. python-pptx `replace_data` обновляет оба ([API chart](https://python-pptx.readthedocs.io/en/latest/api/chart.html)).
- **SmartArt.** Состоит из частей data (dataModel), layout, quickStyle, colors и drawing. В `drawing#.xml` лежит отрендеренный результат. Это расширение Microsoft вне стандарта ([Apache POI XSLFDiagram](https://poi.apache.org/apidocs/dev/org/apache/poi/xslf/usermodel/XSLFDiagram.html); [MS: custom SmartArt layouts](https://msdn.microsoft.com/en-us/library/gg583880.aspx)). При добавлении узла в data.xml PowerPoint перегенерирует drawing при открытии (**гипотеза**: так обычно ведёт себя PowerPoint, но для надёжности ADE должен удалять или помечать устаревший drawing и проверять результат).
- **`normAutofit`.**
  - `fontScale` задаёт процент от исходного размера шрифта.
  - `lnSpcReduction` задаёт процент уменьшения межстрочного интервала.
  - Пример: `<a:normAutofit fontScale="92.000%" lnSpcReduction="20.000%"/>` ([Open XML SDK NormalAutoFit](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc884790(v=office.14))).
  - Это сохранённый результат расчёта, а не правило. ADE должен либо считать значения сам с помощью своего text-measure, либо оставлять расчёт PowerPoint.
- **Placeholders, layouts, masters.** Наследование позиции и стиля идёт по цепочке slide → layout → master. При дублировании плейсхолдер нельзя просто скопировать как фигуру: нужен `idx`/`type`, иначе связь с layout рвётся (**гипотеза**: следует из модели наследования). В python-pptx плейсхолдеры доступны через `slide.placeholders`.

### A.6 Матрица PPTX

| Операция | python-pptx | Office.js | Aspose.Slides | Как / примечание |
|---|---|---|---|---|
| Замена текста с сохранением форматирования run'ов | да | да (1.4 TextRange) | да | Менять текст `a:t` внутри существующих `a:r` |
| Добавить строку таблицы в стиле | частично (XML deepcopy `a:tr`) | да (1.9 `rows.add`) | да (`add_clone`) | [#86](https://github.com/scanny/python-pptx/issues/86), [TableRowCollection](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablerowcollection), [Aspose rows](https://docs.aspose.com/slides/python-net/manage-rows-and-columns/) |
| Дублировать группу (карточку) | частично (XML deepcopy) | частично (`addGroup` есть, clone нет; **гипотеза**: копирование через insertSlides) | да (клонирование фигур, **гипотеза**: `AddClone` для shapes) | [API shapes](https://python-pptx.readthedocs.io/en/latest/api/shapes.html) |
| Перераспределить фигуры | да (left/top/width/height) | да (1.4) | да | Вычисление раскладки за ADE |
| Изменить число категорий диаграммы | да (`replace_data`) | нет | да (`ChartDataWorkbook`) | [chart](https://python-pptx.readthedocs.io/en/latest/api/chart.html), [Aspose](https://docs.aspose.com/slides/python-net/chart-workbook/) |
| Autofit (расчёт) | частично (`fit_text`, грубо) | да (PowerPoint считает сам) | **гипотеза** | [fit_text](https://python-pptx.readthedocs.io/en/latest/dev/analysis/txt-fit-text.html) |
| SmartArt 4 → 5 | нет (только сырой XML) | нет | да (`AddNode`) | [#83](https://github.com/scanny/python-pptx/issues/83), [Aspose nodes](https://docs.aspose.com/slides/net/manage-smartart-shape-node/) |
| Процесс из фигур и коннекторов 4 → 5 | частично (клон + перерасчёт `stCxn`/`endCxn`) | частично (`addLine` 1.4) | да | Коннекторы надо перепривязывать |

---

## B. Google Slides API

### B.1 Операции batchUpdate

В списке Request 49 операций, среди них `insertTableRows`, `duplicateObject`, `groupObjects`/`ungroupObjects`, `updatePageElementTransform`, `replaceAllText`, `createSheetsChart`, `refreshSheetsChart`, `replaceAllShapesWithSheetsChart`, `updateShapeProperties`, `mergeTableCells`, `rerouteLine` ([Request reference](https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request?hl=es-419)).

- **`insertTableRows`.** Вставляет не больше 20 строк за один запрос (там же).
- **`duplicateObject`.** Копия кладётся на ту же страницу в ту же позицию. Слайд-копия вставляется сразу после исходного. `objectIds` задаёт маппинг старых ID на новые ([DuplicateObjectRequest](https://googleapis.dev/ruby/google-apis-slides_v1/v0.22.0/Google/Apis/SlidesV1/DuplicateObjectRequest.html)). Карточку можно продублировать, а затем сдвинуть через `updatePageElementTransform` (RELATIVE или ABSOLUTE). Трансформ группы меняет абсолютные трансформы детей ([Request reference](https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request?hl=es-419)).
- **`groupObjects`.** Нужно минимум два элемента на одной странице, не входящих в другие группы. **Видео, таблицы и плейсхолдеры группировать нельзя** ([GroupObjectsRequest](https://googleapis.dev/ruby/google-apis-slides_v1/v0.12.0/Google/Apis/SlidesV1/GroupObjectsRequest.html)).
- **`replaceAllText`.** Можно ограничить конкретными страницами. Официальный шаблонный подход такой: скопировать шаблон через Drive API, заменить `{{теги}}` через `replaceAllText`, картинки вставить через `replaceAllShapesWithImage` ([Merge guide](https://developers.google.com/workspace/slides/api/guides/merge)).

### B.2 Диаграммы

- Диаграммы бывают **только из Google Sheets**: сначала диаграмма создаётся в Sheets, затем вставляется через `CreateSheetsChartRequest` ([Add chart guide](https://developers.google.com/workspace/slides/api/guides/add-chart?authuser=3)).
- Режимы вставки: `LINKED` (обновляется через `RefreshSheetsChartRequest`) и `NOT_LINKED_IMAGE` (статичная картинка).
- Нативной диаграммы внутри Slides, как в PPTX, **нет**. Чтобы изменить число категорий, нужно поменять диапазон или данные в Sheets через Sheets API и вызвать refresh.
- Даже `LINKED`-диаграмма на слайде остаётся изображением, которое обновляется из Sheets. Редактировать её можно только в Sheets. Строго это противоречит требованию «без картинок» (**гипотеза** о внутреннем представлении, но по поведению это так).

### B.3 Автоподгонка текста

- `Autofit.autofitType` принимает значения NONE, TEXT_AUTOFIT, SHAPE_AUTOFIT.
- `fontScale` и `lineSpacingReduction` **только для чтения**.
- autofitType **автоматически сбрасывается в NONE** любым запросом, который может повлиять на размещение текста. При этом `fontScale` «впекается» в font_size ([Autofit](https://googleapis.dev/ruby/google-apis-slides_v1/v0.12.0/Google/Apis/SlidesV1/Autofit.html)).
- Следствие: после программной правки текста автоподгонку должен считать сам ADE, то есть измерять текст и выставлять размер шрифта или блока.
- Можно ли включить TEXT_AUTOFIT через API: **гипотеза**, что нет или эффект ненадёжен. В найденной документации подтверждения нет.

### B.4 Квоты

Источник: [Usage limits](https://developers.google.com/workspace/slides/limits?authuser=2).
- Чтение: 3000 запросов в минуту на проект, 600 в минуту на пользователя.
- Запись: 600 в минуту на проект, **60 в минуту на пользователя**.
- `getThumbnail`: 300 в минуту на проект, 60 в минуту на пользователя.
- Превышение возвращает 429, нужен экспоненциальный backoff.
- В документации сказано, что позже в 2026 году превышение квот планируют тарифицировать.

Практика: собирать все правки слайда в один `batchUpdate`. Он атомарный, а квота считается по запросам, а не по под-операциям (**гипотеза** о подсчёте).

### B.5 Импорт PPTX в Slides

- Конвертация выполняется через Drive API: загрузить файл с `mimeType` Google Slides. Поддерживаются PowerPoint и ODP ([Drive uploads](https://developers.google.com/workspace/drive/api/guides/manage-uploads)).
- Потери:
  - **все диаграммы конвертируются в изображения** ([think-cell KB0224](https://www.thinkcell.com/en/resources/kb/0224));
  - встречаются искажения подписей и осей, скрытые фигуры удаляются (там же);
  - SmartArt превращается в фигуры или изображение (**гипотеза**: официального перечня Google не найдено, в справке Google ([Switch from PowerPoint](https://support.google.com/docs/answer/9310378?hl=en)) этого нет).

### B.6 Матрица Google Slides

| Операция | Поддержка | Как | Источник |
|---|---|---|---|
| Замена текста | да | `replaceAllText`, `deleteText` + `insertText` | [Merge](https://developers.google.com/workspace/slides/api/guides/merge) |
| Добавить строку таблицы в стиле | да | `insertTableRows` (`insertBelow`, наследует стиль соседней строки; **гипотеза** о наследовании) | [Request](https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request?hl=es-419) |
| Дублировать карточку | да | `duplicateObject` для группы + `updatePageElementTransform` | [Duplicate](https://googleapis.dev/ruby/google-apis-slides_v1/v0.22.0/Google/Apis/SlidesV1/DuplicateObjectRequest.html) |
| Группировать | частично | `groupObjects`, без таблиц и плейсхолдеров | [GroupObjects](https://googleapis.dev/ruby/google-apis-slides_v1/v0.12.0/Google/Apis/SlidesV1/GroupObjectsRequest.html) |
| Изменить число категорий диаграммы | частично | Только через Sheets + `refreshSheetsChart` | [Add chart](https://developers.google.com/workspace/slides/api/guides/add-chart?authuser=3) |
| Нативная диаграмма | нет | — | там же |
| Autofit | нет (только чтение) | Считать самим | [Autofit](https://googleapis.dev/ruby/google-apis-slides_v1/v0.12.0/Google/Apis/SlidesV1/Autofit.html) |
| Процесс 4 → 5 | частично | `duplicateObject` + transform; линии через `rerouteLine` (**гипотеза**: работает для коннекторов) | [Request](https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request?hl=es-419) |

---

## C. Figma / Figma Slides

### C.1 Какие API есть

- **REST API** «largely read-only». Писать он может только Comments, Comment Reactions, Variables и Dev Resources ([Compare APIs](https://developers.figma.com/compare-apis/)). Изменять узлы через REST **нельзя**.
- **Plugin API** работает внутри открытого файла и позволяет «editing almost any aspect» (там же).
- **MCP server.** `use_figma` (только remote) создаёт, редактирует, удаляет и инспектирует объекты в **Figma Design, FigJam и Figma Slides**. `create_new_file` создаёт Design, FigJam или Slides ([MCP tools](https://developers.figma.com/docs/figma-mcp-server/tools-and-prompts/)).
  - Запись на холст: «free during the beta period», затем станет платной по использованию ([MCP guide README](https://cdn.jsdelivr.net/gh/figma/mcp-server-guide@main/README.md)).
  - Через skill `figma-use-slides` можно работать со Slides-файлами при наличии edit-доступа ([Figma skills for MCP](https://help.figma.com/hc/en-us/articles/39166810751895-Figma-skills-for-MCP)).
- **Лимиты MCP** ([rate limits](https://developers.figma.com/docs/figma-mcp-server/rate-limits-access/)):
  - View и Collab seat: 20 вызовов в месяц;
  - Dev/Full на Professional: 200 в день и 10 в минуту;
  - Organization: 200 в день и 15 в минуту;
  - Enterprise: 600 в день и 20 в минуту.
  - Пишущие инструменты частично освобождены от лимитов, но `use_figma` в списке освобождённых **явно не назван**.
- **Лимиты REST** ([rate limits](https://developers.figma.com/docs/rest-api/rate-limits)):
  - Tier 1 (GET file): 10, 15 или 20 запросов в минуту для Dev/Full в зависимости от плана;
  - для View/Collab: до 20 в месяц.

### C.2 Auto-layout и компоненты

- `layoutMode` принимает значения NONE, HORIZONTAL, VERTICAL, GRID. Изменение режима перемещает детей. Отключение auto-layout не возвращает исходные позиции ([layoutMode](https://developers.figma.com/docs/plugins/api/properties/nodes-layoutmode/)).
- С обновления 140 (30.09.2026) `layoutWrap = WRAP` работает и для вертикального auto-layout ([update 140](https://developers.figma.com/docs/plugins/updates/2026/09/30/version-1-update-140/)).
- Для ADE «добавить карточку» в auto-layout фрейме сводится к `clone()` и `insertChild`, раскладку пересчитывает Figma. Это главное преимущество Figma перед PPTX и Slides, где перераспределение считает сам ADE.
- Для текста есть `textAutoResize` (NONE, WIDTH_AND_HEIGHT, HEIGHT, TRUNCATE), `textTruncation` и `maxLines` ([TextNode](https://developers.figma.com/docs/plugins/api/TextNode)). Перед изменением текста нужно загрузить шрифт ([node properties](https://developers.figma.com/docs/plugins/api/node-properties/)). Уменьшения шрифта «под рамку», как в PowerPoint, нет: подгонку размера шрифта считает ADE (**гипотеза**: такого режима среди перечисленных значений нет).

### C.3 Таблицы и диаграммы

- **Таблицы.** `createTable()` есть в Slides (и в FigJam), но **не в Figma Design**: раздел «Supported in Slides but not in Design» ([Working in Slides](https://developers.figma.com/docs/plugins/working-in-slides)). У `TableNode` есть `insertRow` ([TableNode](https://developers.figma.com/docs/plugins/api/TableNode)). В UI Slides таблицы ограничены 500 ячейками ([Add tables to slides](https://help.figma.com/hc/en-us/articles/30600895164439-Add-tables-to-slides)).
- **Figma Slides не поддерживает Components, Styles, Variables и Libraries** в Plugin API ([Working in Slides](https://developers.figma.com/docs/plugins/working-in-slides)). Значит, подход «карточка как инстанс компонента» в Slides недоступен, хотя auto-layout фреймы остаются (**гипотеза**: auto-layout в Slides доступен через API, явно на странице не оговорено). Размер слайда фиксирован: 1920×1080.
- **Диаграммы.** Нативных диаграмм нет. Feature request в форуме с 2024 года без релиза ([forum](https://forum.figma.com/suggest-a-feature-11/adding-simple-chart-to-figma-slides-like-powerpoint-15597)). В обзоре Config 2026 диаграмм тоже нет ([classmethod recap](https://dev.classmethod.jp/articles/figma-config-2026-recap/)). Диаграмма в Figma — это векторные фигуры или результат плагина, то есть «редактируемая геометрия», а не данные.

### C.4 Импорт и экспорт PPTX в Figma Slides

- **Импорт** ([Import PPTX](https://help.figma.com/hc/hi/articles/30601628883607)):
  - таблицы, диаграммы (diagrams), анимации и видео **не импортируются**;
  - все шрифты заменяются на **Inter**.
- **Экспорт в PPTX** ([Export](https://help.figma.com/hc/en-us/articles/24848334599447-Export-from-Figma-Slides)):
  - можно выбрать «editable PPTX equivalent» или растр;
  - градиенты становятся сплошной заливкой;
  - интерактивные элементы становятся картинками;
  - отсутствующие в PowerPoint шрифты заменяются шрифтом по умолчанию.

### C.5 Матрица Figma

| Операция | Поддержка | Как | Источник |
|---|---|---|---|
| Замена текста | да (Plugin/MCP) | `characters` после `loadFontAsync` | [node props](https://developers.figma.com/docs/plugins/api/node-properties/) |
| Замена текста через REST | нет | — | [Compare APIs](https://developers.figma.com/compare-apis/) |
| Добавить строку таблицы | частично (Slides/FigJam) | `TableNode.insertRow`; в Design таблиц нет, там «таблица» — это фреймы | [TableNode](https://developers.figma.com/docs/plugins/api/TableNode), [Slides](https://developers.figma.com/docs/plugins/working-in-slides) |
| Добавить карточку | да (лучший случай) | `clone()` в auto-layout фрейме; в Design ещё и инстанс компонента | [layoutMode](https://developers.figma.com/docs/plugins/api/properties/nodes-layoutmode/) |
| Изменить число категорий диаграммы | нет (нативных нет) | Перестроить вектор | [forum](https://forum.figma.com/suggest-a-feature-11/adding-simple-chart-to-figma-slides-like-powerpoint-15597) |
| Autofit | частично | `textAutoResize`; уменьшение шрифта считает ADE | [TextNode](https://developers.figma.com/docs/plugins/api/TextNode) |
| Процесс 4 → 5 | да | Auto-layout ряд + clone; коннекторы в Slides недоступны (`createConnector` недоступен) | [Slides](https://developers.figma.com/docs/plugins/working-in-slides) |

---

## D. Различия объектных моделей и что это значит для IDM

| Аспект | PPTX | Google Slides | Figma |
|---|---|---|---|
| Координаты | Абсолютные, EMU; группа задаёт свою систему координат (`chOff`/`chExt`) | Абсолютные, affine transform (EMU/PT) | Относительно родителя, плюс auto-layout поток |
| Раскладка | Нет движка; всё абсолютно | Нет движка | **Есть** движок: auto-layout и constraints |
| Наследование стиля | master → layout → placeholder; темы | layouts и placeholders, наследование autofit от плейсхолдера | компоненты и инстансы, styles и variables (не в Slides) |
| Таблица | Нативная (`a:tbl`) | Нативная | Только Slides/FigJam |
| Диаграмма | Нативная с данными (xlsx + кэш) | Только Sheets-linked | Нет |
| Схема/SmartArt | SmartArt (data + layout + drawing) | Нет (фигуры) | Нет (фигуры) |
| Autofit | Хранится результат (`fontScale`) | Только чтение, сбрасывается | `textAutoResize` без уменьшения шрифта |
| Идентификаторы | `cNvPr id` (на слайде) | objectId (строка, задаётся клиентом) | node id |

Выводы для IDM:
1. **IDM должна быть семантической, а не геометрической.** Нужны узлы `Repeater` (карточки, этапы процесса, строки таблицы) с шаблоном элемента и правилом распределения (ось, gap, выравнивание), плюс `DataChart` (категории, серии, тип), `Table` и `TextSlot` (бокс, min/max шрифта, политика overflow). Геометрию на каждой платформе считает собственный солвер ADE. В Figma её можно делегировать auto-layout.
2. **Возможности задаются по адаптерам.** Каждый адаптер объявляет, какие операции поддерживает: native, emulated или unsupported. Например, `DataChart` в Figma может быть только `unsupported` или «векторная перестройка», а это нарушает требование редактируемости данных. Это нужно показывать пользователю явно.
3. **Измерение текста — общий компонент ADE.** Ни один API не даёт надёжного «shrink to fit» после программной правки: python-pptx не считает, в Google значения только читаются и сбрасываются, в Figma такого режима нет. Нужен собственный text-measure (HarfBuzz/FreeType + фирменные шрифты) и проверка рендером.
4. **Шаблонные цепочки PPTX (master, layout, тема) надо сохранять в IDM как ссылки**, а не разворачивать в «плоские» стили. Иначе пострадает редактируемость, например смена темы.
5. **Round-trip через импорт (PPTX → Slides или Figma) неприемлем** как основной путь: диаграммы превращаются в картинки или теряются, шрифты заменяются. Адаптеры должны работать с нативными файлами каждой платформы.
6. **Стабильные ID.** В Slides objectId задаётся клиентом при дублировании (`objectIds` map), это удобно для идемпотентности. В PPTX ID нужно перенумеровывать вручную.

---

## E. Открытые проекты и исследования, близкие к ADE

| Проект | Что делает | Близость к ADE | Источник |
|---|---|---|---|
| **PPTAgent** (arXiv 2501.03936, 2025) | Анализирует референсные презентации (функциональные типы слайдов, content schema). Затем **edit-based** генерация: API `del_span`, `del_image`, `clone_paragraph`, `replace_span`, `replace_image` поверх HTML-представления слайда, до двух итераций самокоррекции. Оценка через PPTEval (Content, Design, Coherence) | Высокая: правки существующих слайдов вместо генерации с нуля. Но нет таблиц-строк, диаграмм и перераспределения | [arXiv abs](https://arxiv.org/abs/2501.03936v3), [HTML](https://arxiv.org/html/2501.03936v3), [GitHub](https://github.com/icip-cas/PPTAgent) |
| **Talk to Your Slides** (arXiv 2505.11604, 2025) | Агент правит слайды через структурированные данные, а не через GUI. Бенчмарк TSBench (379 инструкций). Заявлено на 34% быстрее и на 87% дешевле GUI-агентов | Средняя: редактирование, но фокус на тексте и форматировании | [arXiv](https://arxiv.org/abs/2505.11604v3) |
| **Presenton** (Apache 2.0) | Делает шаблоны из существующего PPTX: layout в HTML + Tailwind, затем TSX + Zod-схема. Экспорт в PPTX и PDF, заявлен как «editable» | Средняя: шаблонизация, но через HTML, а не нативные объекты | [GitHub](https://github.com/presenton/presenton), [docs](https://docs.presenton.ai/create-presentation-template-with-ai.md) |
| **SlideDeck AI** | LLM генерирует JSON по схеме, python-pptx собирает слайды по трём фиксированным шаблонам | Низкая | [README](https://huggingface.co/spaces/barunsaha/slide-deck-ai/blob/9c0dccddf77db6ebf7934a4979eb4a301eca104b/README.md) |
| **pptx-renderer** | Шаблонизатор PPTX: `{{{expr}}}`, вставка `table()` и `image()` в прямоугольник, `loop_groups` для повтора слайдов | Средняя: ближайший аналог docxtpl, но таблицы создаёт заново, а не продолжает стиль | [docs](https://pptx-renderer.readthedocs.io/en/master/usage.html) |
| **python-pptx-templater** | Jinja + layouts, последний релиз 2019 года | Низкая (заброшен) | [PyPI](https://pypi.org/project/python-pptx-templater/) |
| **Google Slides merge** (официальный паттерн) | Копия шаблона, `replaceAllText`, `replaceAllShapesWithImage` | Базовый слой замены текста | [guide](https://developers.google.com/workspace/slides/api/guides/merge) |
| **Adaptive Grid-Based Document Layout** (Jacobs et al., SIGGRAPH 2003) | Шаблоны на сетке с ограничениями адаптируются к контенту, страницы подбирает оптимизатор | Теоретическая основа солвера раскладки ADE | [Adobe Research](https://research.adobe.com/publication/adaptive-grid-based-document-layout) |
| **DocLayNet** (IBM, 2022) | Датасет разметки макетов документов | Для обучения и оценки распознавания ролей блоков, а не для адаптации | [arXiv 2206.01062](https://arxiv.org/abs/2206.01062v1), [HF](https://huggingface.co/datasets/docling-project/DocLayNet-v1.1/blob/main/README.md) |

Пробел на рынке (**гипотеза**, основанная на обзоре выше): открытого инструмента, который *продолжает* существующий макет (строки, карточки, категории, этапы) нативными объектами с перераспределением геометрии, не найдено. PPTAgent ближе всех по идее (правка референса), но работает на уровне текстовых span'ов.

---

## F. Что проверить экспериментально (не подтверждено документацией)

1. Перегенерирует ли PowerPoint (Win/Mac/Web) `drawing.xml` SmartArt после правки `data.xml`, если старый drawing удалить.
2. Пересчитывает ли PowerPoint `normAutofit fontScale` при открытии, без редактирования (важно для пакетной генерации).
3. Точность рендера LibreOffice и Aspose против PowerPoint: метрика вроде SSIM или IoU bbox текста на корпусе из 50 слайдов с фирменными шрифтами.
4. Наследует ли `insertTableRows` в Google Slides стиль ячеек соседней строки полностью: заливку, границы, стиль текста.
5. Доступны ли auto-layout фреймы в Figma Slides через `use_figma` и работает ли `TableNode.insertRow` через MCP.
6. Учитывается ли `use_figma` в лимитах MCP (200 вызовов в день на Professional), что критично для пакетной адаптации.
