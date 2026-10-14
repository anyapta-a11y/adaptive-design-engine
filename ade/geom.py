"""Геометрия и идентификаторы shape-элементов (lxml поверх python-pptx)."""
from __future__ import annotations

from pptx.oxml.ns import qn

SP, GRP, CXN, GF, PIC = (qn("p:sp"), qn("p:grpSp"), qn("p:cxnSp"),
                         qn("p:graphicFrame"), qn("p:pic"))
SHAPE_TAGS = (SP, GRP, CXN, GF, PIC)
EMU_PT = 12700


def xfrm(el):
    if el.tag == GF:
        return el.find(qn("p:xfrm"))
    pr = el.find(qn("p:grpSpPr")) if el.tag == GRP else el.find(qn("p:spPr"))
    return None if pr is None else pr.find(qn("a:xfrm"))


def _pair(x, a, b):
    e = x.find(qn(a))
    return None if e is None else (int(e.get(b[0])), int(e.get(b[1])))


def get_box(el):
    x = xfrm(el)
    if x is None:
        return None
    off, ext = _pair(x, "a:off", "xy"), _pair(x, "a:ext", ("cx", "cy"))
    if off is None or ext is None:
        return None
    return (off[0], off[1], ext[0], ext[1])


def set_box(el, x, y, cx, cy):
    xf = xfrm(el)
    o, e = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    o.set("x", str(int(round(x)))); o.set("y", str(int(round(y))))
    e.set("cx", str(int(round(cx)))); e.set("cy", str(int(round(cy))))


def set_group_box(el, x, y, cx, cy):
    """Двигает/масштабирует элемент; у группы пересчитывает детей и 'запекает' chOff/chExt."""
    if el.tag != GRP:
        set_box(el, x, y, cx, cy)
        return
    xf = xfrm(el)
    ch_o, ch_e = xf.find(qn("a:chOff")), xf.find(qn("a:chExt"))
    cox, coy = int(ch_o.get("x")), int(ch_o.get("y"))
    cex, cey = int(ch_e.get("cx")) or 1, int(ch_e.get("cy")) or 1
    sx, sy = cx / cex, cy / cey
    for c in el:
        if c.tag in SHAPE_TAGS and get_box(c):
            bx, by, bw, bh = get_box(c)
            set_box(c, x + (bx - cox) * sx, y + (by - coy) * sy, bw * sx, bh * sy)
    set_box(el, x, y, cx, cy)
    ch_o.set("x", str(int(round(x)))); ch_o.set("y", str(int(round(y))))
    ch_e.set("cx", str(int(round(cx)))); ch_e.set("cy", str(int(round(cy))))


def name_of(el):
    c = el.find("./*/" + qn("p:cNvPr"))
    return c.get("name") if c is not None else None


def set_name(el, name):
    el.find("./*/" + qn("p:cNvPr")).set("name", name)


def walk(elements, tf=(1.0, 1.0, 0.0, 0.0), parent=None):
    """Рекурсивный обход: (элемент, абсолютный бокс EMU, родительская группа)."""
    for el in elements:
        if el.tag not in SHAPE_TAGS:
            continue
        b = get_box(el)
        if b is None:
            continue
        ab = (b[0] * tf[0] + tf[2], b[1] * tf[1] + tf[3], b[2] * tf[0], b[3] * tf[1])
        yield el, ab, parent
        if el.tag == GRP:
            xf = xfrm(el)
            ch_o, ch_e = _pair(xf, "a:chOff", "xy"), _pair(xf, "a:chExt", ("cx", "cy"))
            sx = b[2] / ch_e[0] if ch_e[0] else 1.0
            sy = b[3] / ch_e[1] if ch_e[1] else 1.0
            ntf = (tf[0] * sx, tf[1] * sy,
                   ab[0] - ch_o[0] * tf[0] * sx, ab[1] - ch_o[1] * tf[1] * sy)
            yield from walk(list(el), ntf, el)


def ids(root):
    return [int(c.get("id")) for c in root.iter(qn("p:cNvPr"))]


def renumber(el, slide_root):
    """Присваивает всем cNvPr клона новые id (max на слайде + 1...)."""
    nxt = max(ids(slide_root) or [0]) + 1
    for c in el.iter(qn("p:cNvPr")):
        c.set("id", str(nxt)); nxt += 1
