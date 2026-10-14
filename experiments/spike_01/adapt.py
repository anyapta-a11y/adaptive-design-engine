"""Spike 01: проверка пяти структурных адаптаций на python-pptx + прямой OOXML.
Не продуктовый код: цель — доказать или опровергнуть техническую реализуемость."""
import copy, sys
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.chart.data import CategoryChartData
from PIL import ImageFont

FONT_PATH = "/usr/share/fonts/opentype/inter/Inter-Regular.otf"
FONT_BOLD_PATH = "/usr/share/fonts/opentype/inter/Inter-Bold.otf"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

def _max_id(slide):
    return max([int(i) for i in slide._element.xpath('.//p:cNvPr/@id')] or [0])

def clone_after(slide, src_el):
    """Клонирует элемент после оригинала; все cNvPr клона получают новые id (max на слайде + 1..)."""
    nxt = _max_id(slide) + 1
    el = copy.deepcopy(src_el)
    for c in el.xpath('.//p:cNvPr'):
        c.set("id", str(nxt)); nxt += 1
    src_el.addnext(el)
    return el

def _insets(shape):
    bp = shape.text_frame._txBody.find(A_NS + "bodyPr")
    l = int(bp.get("lIns", 91440)) if bp is not None else 91440
    r = int(bp.get("rIns", 91440)) if bp is not None else 91440
    return l + r

def _longest_word_px(shape, pt):
    """Ширина самого длинного слова при pt (в единицах pt*10) и признак: (px, avail_px)."""
    worst = 0
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            path = FONT_BOLD_PATH if r.font.bold else FONT_PATH
            f = ImageFont.truetype(path, size=max(1, int(round(pt * 10))))  # размер None = умолчание 18pt (как в qa.py)
            for w in r.text.split():
                worst = max(worst, f.getlength(w))
    return worst

def fit_uniform(shapes, kind):
    """Единый кегль для однотипных элементов: шаг 1pt вниз, не ниже 75% исходного.
    Текст не сокращается; если и на 75% не влезает - предупреждение в stdout."""
    shapes = [s for s in shapes if s.has_text_frame and s.text_frame.text.strip()]
    runs = [r for s in shapes for p in s.text_frame.paragraphs for r in p.runs]
    if not runs: return None
    orig = min(r.font.size.pt if r.font.size else 18.0 for r in runs); floor = orig * 0.75
    def fits(pt):
        return all(_longest_word_px(s, pt) <= (s.width - _insets(s)) / 12700 * 10 for s in shapes)
    pt = orig
    while not fits(pt) and pt - 1 >= floor: pt -= 1
    if not fits(pt):
        pt = max(floor, pt)   # минимум 75%; финальный размер = пол
        print(f"WARNING: {kind}: слово не помещается даже при {pt:g}pt (75% от {orig:g}pt); текст не сокращён")
    if pt != orig:
        for s in shapes:
            for p in s.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(pt)
    return pt

def by_name(slide, name):
    return next(s for s in slide.shapes if s.name == name)

def set_text_keep_style(shape, text):
    """Заменяет текст в первом run, сохраняя его форматирование."""
    p = shape.text_frame.paragraphs[0]
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)

# ---------- 1. Карточки: N -> M с перераспределением ----------
def _move_group(grp, new_left, new_width):
    old_left, old_width = grp.left, grp.width
    sx = new_width / old_width
    for ch in grp.shapes:
        ch.left = Emu(int(new_left + (ch.left - old_left) * sx))
        ch.width = Emu(int(ch.width * sx))
    xfrm = grp._element.grpSpPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
    for tag in ("off", "chOff"):
        xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}' + tag).set("x", str(int(new_left)))
    for tag in ("ext", "chExt"):
        xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}' + tag).set("cx", str(int(new_width)))

def adapt_cards(slide, items):
    cards = sorted([s for s in slide.shapes if s.name.startswith("Card ")], key=lambda s: s.left)
    left, right = cards[0].left, cards[-1].left + cards[-1].width
    gap = cards[1].left - (cards[0].left + cards[0].width)
    while len(cards) < len(items):   # клонируем последнюю карточку целиком (стиль сохраняется)
        clone_after(slide, cards[-1]._element)
        cards = sorted([s for s in slide.shapes if s.name.startswith("Card ")], key=lambda s: s.left)
        cards.append(cards.pop(-1))  # клон стоит последним
    n = len(items)
    w = (right - left - gap * (n - 1)) / n
    for i, (card, (t, body)) in enumerate(zip(cards, items)):
        _move_group(card, left + i * (w + gap), w)
        card.name = f"Card {i+1}"
        set_text_keep_style(next(c for c in card.shapes if c.name == "Card title"), t)
        set_text_keep_style(next(c for c in card.shapes if c.name == "Card body"), body)
    for nm in ("Card title", "Card body"):
        fit_uniform([c for card in cards for c in card.shapes if c.name == nm], nm)

# ---------- 2. Таблица: добавление строк с сохранением чередования ----------
def adapt_table(slide, rows, max_bottom):
    gf = by_name(slide, "Data table"); tbl = gf.table._tbl
    trs = tbl.tr_lst
    while len(trs) < len(rows):
        src = trs[-2] if len(trs) >= 3 else trs[-1]   # шаблон через одну строку = чередование заливки
        new = copy.deepcopy(src); trs[-1].addnext(new); trs = tbl.tr_lst
    while len(trs) > len(rows):
        tbl.remove(trs[-1]); trs = tbl.tr_lst
    for r_i, row in enumerate(rows):
        for c_i, v in enumerate(row):
            cell = gf.table.cell(r_i, c_i)
            set_text_keep_style(cell, v)
    # высота: сохраняем исходную высоту строки, если влезает; иначе сжимаем равномерно
    row_h = gf.table.rows[0].height
    total = row_h * len(rows)
    if gf.top + total > max_bottom:
        row_h = int((max_bottom - gf.top) / len(rows)); total = row_h * len(rows)
    for r in gf.table.rows: r.height = Emu(row_h)
    gf.height = Emu(total)

# ---------- 3. Диаграмма: новые категории в нативной диаграмме ----------
def adapt_chart(slide, cats, series):
    ch = by_name(slide, "Market chart").chart
    cd = CategoryChartData(); cd.categories = cats
    for name, vals in series: cd.add_series(name, vals)
    ch.replace_data(cd)   # обновляет и XML-кэш, и встроенный xlsx

# ---------- 4. Процесс: N -> M этапов в том же пролёте ----------
def adapt_process(slide, labels):
    circles = sorted([s for s in slide.shapes if s.name.endswith("circle")], key=lambda s: s.left)
    lbls = sorted([s for s in slide.shapes if s.name.endswith("label")], key=lambda s: s.left)
    conns = sorted([s for s in slide.shapes if s.name.startswith("Connector")], key=lambda s: s.left)
    span_l, span_r = circles[0].left, circles[-1].left + circles[-1].width
    while len(circles) < len(labels):
        for coll in (circles, lbls, conns):
            clone_after(slide, coll[-1]._element)
        circles = sorted([s for s in slide.shapes if s.name.endswith("circle")], key=lambda s: s._element.getparent().index(s._element))
        lbls = sorted([s for s in slide.shapes if s.name.endswith("label")], key=lambda s: s._element.getparent().index(s._element))
        conns = sorted([s for s in slide.shapes if s.name.startswith("Connector")], key=lambda s: s._element.getparent().index(s._element))
    n = len(labels); d = circles[0].width
    min_gap = int(d * 0.35)
    if (span_r - span_l - n * d) / (n - 1) < min_gap:      # уменьшаем элементы, сохраняя пропорции
        d = int((span_r - span_l - min_gap * (n - 1)) / n)
    step = (span_r - span_l - d) / (n - 1)
    lw = lbls[0].width
    for i in range(n):
        c, l = circles[i], lbls[i]
        cy_mid = c.top + c.height // 2
        c.left, c.width, c.height, c.top = Emu(int(span_l + i * step)), Emu(d), Emu(d), Emu(cy_mid - d // 2)
        c.name = f"Step {i+1} circle"; set_text_keep_style(c, str(i + 1))
        l.width = Emu(int(min(lw, step - 50000)))
        l.left = Emu(int(c.left + d / 2 - l.width / 2)); l.name = f"Step {i+1} label"
        set_text_keep_style(l, labels[i])
    for i, cn in enumerate(conns):
        x1 = circles[i].left + d; x2 = circles[i + 1].left; y = circles[i].top + d // 2
        cn.begin_x, cn.begin_y, cn.end_x, cn.end_y = Emu(int(x1)), Emu(int(y)), Emu(int(x2)), Emu(int(y))
        cn.name = f"Connector {i+1}"
    fit_uniform(lbls[:n], "Step label")

# ---------- 5. Текст: измерение и перераспределение, без сокращения текста ----------
def text_height_emu(text, width_emu, pt, line_spacing=1.2, inset_emu=91440*2):
    font = ImageFont.truetype(FONT_PATH, size=pt * 10)   # 10x для точности; 1pt = 12700 EMU
    width_px = (width_emu - inset_emu) / 12700 * 10
    lines, cur = 0, ""
    for para in text.split("\n"):
        cur = ""; lines += 1
        for w in para.split(" "):
            cand = (cur + " " + w).strip()
            if font.getlength(cand) <= width_px: cur = cand
            else: lines += 1; cur = w
    return int(lines * pt * line_spacing * 12700 + 91440), lines

def adapt_text(slide, body, max_bottom, min_pt=14):
    box = by_name(slide, "Body text"); foot = by_name(slide, "Footnote")
    pt = box.text_frame.paragraphs[0].runs[0].font.size.pt
    set_text_keep_style(box, body)
    gap = foot.top - (box.top + box.height)
    while True:
        h, lines = text_height_emu(body, box.width, pt)
        if box.top + h + gap + foot.height <= max_bottom or pt <= min_pt: break
        pt -= 1
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(pt)
    box.height = Emu(max(box.height, h)); foot.top = Emu(box.top + box.height + gap)
    return {"font_pt": pt, "lines": lines, "fits": foot.top + foot.height <= max_bottom}

if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    prs = Presentation(src); S = prs.slides; bottom = prs.slide_height - 457200
    adapt_cards(S[0], [("Скорость", "Подготовка презентации за минуты вместо дней."),
                       ("Стиль", "Фирменный стиль сохраняется без ручной правки."),
                       ("Контроль", "Каждое изменение проверяется автоматически."),
                       ("Редактируемость", "Все объекты остаются нативными объектами PowerPoint.")])
    adapt_table(S[1], [("Регион", "Выручка", "Рост")] + [(f"Регион {i}", f"{200+i*7}", f"{i}%") for i in range(1, 7)], bottom)
    adapt_chart(S[2], ["A", "B", "C", "D"], [("2026", (28, 40, 20, 12))])
    adapt_process(S[3], ["Аудит", "Пилот", "Обучение", "Масштабирование", "Поддержка"])
    rep = adapt_text(S[4], ("Новый текст резюме существенно длиннее исходного: агент обязан разместить его полностью, "
                            "не сокращая формулировки клиента, сохранив иерархию и не допустив наложения на сноску. ") * 3, bottom)
    prs.save(dst); print("text:", rep)
