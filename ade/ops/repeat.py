"""Повторяемые элементы item[n] / connector[n]: клон, удаление, перераспределение в пролёте."""
from __future__ import annotations

import copy
import re

from pptx.oxml.ns import qn

from .. import geom
from ..errors import AdeError

ITEM_RE = re.compile(r"^item\[(\d+)\]$")
CONN_RE = re.compile(r"^connector\[(\d+)\]$")
DEFAULT_GAP = 274320        # 0.3", если элемент один
MIN_GAP_FACTOR = 0.5        # минимальный зазор при сжатии = 50% исходного
PAD = 54864                 # 0.06" отступ коннектора от опорной фигуры


def _indexed(tree, rx):
    out = {}
    for el in tree:
        if el.tag in geom.SHAPE_TAGS:
            m = rx.match(geom.name_of(el) or "")
            if m:
                out[int(m.group(1))] = el
    return out


def _rename(el, old: str, new: str) -> None:
    for c in el.iter(qn("p:cNvPr")):
        n = c.get("name") or ""
        if n.startswith(old):
            c.set("name", new + n[len(old):])


def _clone_last(tree, els: dict, kind: str):
    k = max(els)
    src = els[k]
    new = copy.deepcopy(src)
    src.addnext(new)
    geom.renumber(new, tree.getroottree().getroot())
    _rename(new, f"{kind}[{k}]", f"{kind}[{k + 1}]")
    els[k + 1] = new
    return new


UNIT_RE = re.compile(r"^item\[(\d+)\](?:\.(\w+))?$")
SAFE = 0.05                 # безопасное поле: доля размера слайда с каждого края


def unit_elements(tree, k: int) -> list:
    """Все фигуры верхнего уровня элемента k: группа/фигура 'item[k]' и/или 'item[k].<поле>'."""
    out = []
    for el in tree:
        if el.tag in geom.SHAPE_TAGS:
            m = UNIT_RE.match(geom.name_of(el) or "")
            if m and int(m.group(1)) == k:
                out.append(el)
    return out


def unit_numbers(tree) -> list:
    return sorted({int(m.group(1)) for el in tree if el.tag in geom.SHAPE_TAGS
                   for m in [UNIT_RE.match(geom.name_of(el) or "")] if m})


def unit_box(els):
    bs = [b for b in (geom.get_box(e) for e in els) if b]
    l, t = min(b[0] for b in bs), min(b[1] for b in bs)
    return (l, t, max(b[0] + b[2] for b in bs) - l, max(b[1] + b[3] for b in bs) - t)


def detect_axis(boxes) -> str:
    """Ось повторения по центрам item[1]..item[n]: больший разброс по y -> 'y', иначе 'x'."""
    if len(boxes) < 2:
        return "x"
    cx = [b[0] + b[2] / 2 for b in boxes]
    cy = [b[1] + b[3] / 2 for b in boxes]
    return "y" if (max(cy) - min(cy)) > (max(cx) - min(cx)) else "x"


def _clone_unit(tree, k: int) -> list:
    src = unit_elements(tree, k)
    cur, out = src[-1], []
    for el in src:
        new = copy.deepcopy(el)
        cur.addnext(new)
        cur = new
        geom.renumber(new, tree.getroottree().getroot())
        _rename(new, f"item[{k}]", f"item[{k + 1}]")
        out.append(new)
    return out


def _shift(els, dx, dy):
    for e in els:
        b = geom.get_box(e)
        if b:
            geom.set_group_box(e, b[0] + dx, b[1] + dy, b[2], b[3])


def _span_limit(tree, uels, a, size, spans_end):
    """Конец пролёта по оси: безопасный край слайда или начало ближайшей фигуры за элементами."""
    c = 1 - a
    allb = [geom.get_box(e) for e in uels]
    lo = min(b[c] for b in allb)
    hi = max(b[c] + b[c + 2] for b in allb)
    limit = size[a] * (1 - SAFE)
    ids = {id(e) for e in uels}
    for el in tree:
        if el.tag not in geom.SHAPE_TAGS or id(el) in ids:
            continue
        b = geom.get_box(el)
        if not b or b[a + 2] > 0.9 * size[a] or b[c + 2] > 0.9 * size[c]:
            continue
        if b[a] >= spans_end - 1 and b[c] < hi and b[c] + b[c + 2] > lo:
            limit = min(limit, b[a])
    return max(limit, spans_end)


def adapt_items(tree, target: int, proportional: bool = False, axis: str | None = None,
                slide_size: tuple | None = None) -> dict:
    """Приводит число item[n] к target.
    Возвращает {'cloned','deleted','scale','ratio_w','ratio_h','axis','warnings'}."""
    nums = unit_numbers(tree)
    n0 = len(nums)
    if n0 == 0:
        raise AdeError("В слайде нет item[n]")
    if target < 1:
        raise AdeError("Число элементов должно быть >= 1")
    units = {k: unit_elements(tree, k) for k in nums}
    boxes = [unit_box(units[k]) for k in nums]
    axis = axis or detect_axis(boxes)
    base = {"cloned": 0, "deleted": 0, "scale": 1.0, "ratio_w": 1.0, "ratio_h": 1.0,
            "axis": axis, "warnings": []}
    if target == n0:
        return base
    conns = _indexed(tree, CONN_RE)
    single = all(len(units[k]) == 1 and geom.name_of(units[k][0]) == f"item[{k}]" for k in nums)
    cloned = deleted = 0
    while len(units) < target:
        k = max(units)
        units[k + 1] = _clone_unit(tree, k)
        if conns:
            _clone_last(tree, conns, "connector")
        cloned += 1
    while len(units) > target:
        for e in units.pop(max(units)):
            tree.remove(e)
        deleted += 1
        if conns:
            tree.remove(conns.pop(max(conns)))
    res = dict(base, cloned=cloned, deleted=deleted)
    if single and axis == "x":
        res.update(_place_legacy(units, conns, boxes, target, n0, proportional))
    else:
        res.update(_place_keep_size(tree, units, conns, boxes, target, n0, axis, slide_size, res))
    nb = [unit_box(units[k]) for k in sorted(units)]
    res["ratio_w"] = (sum(b[2] for b in nb) / len(nb)) / (sum(b[2] for b in boxes) / n0 or 1)
    res["ratio_h"] = (sum(b[3] for b in nb) / len(nb)) / (sum(b[3] for b in boxes) / n0 or 1)
    return res


def _place_legacy(units, conns, boxes, n, n0, proportional):
    """Горизонтальные одиночные item: ширина элементов подстраивается под пролёт."""
    items = {k: v[0] for k, v in units.items()}
    span_l = min(b[0] for b in boxes)
    span_r = max(b[0] + b[2] for b in boxes)
    xs = sorted(boxes)
    gaps = [xs[i + 1][0] - (xs[i][0] + xs[i][2]) for i in range(len(xs) - 1)]
    g0 = max(0, sum(gaps) / len(gaps)) if gaps else DEFAULT_GAP
    w0 = boxes[0][2]
    span = span_r - span_l
    scale = 1.0
    if n == 1:
        w, gap = min(w0, span), 0.0
    elif proportional:
        gmin = g0 * MIN_GAP_FACTOR
        if n * w0 + (n - 1) * gmin <= span:
            w = w0
        else:
            w = (span - (n - 1) * gmin) / n
            scale = w / w0
        gap = (span - n * w) / (n - 1)
    else:
        gap = g0 * (n0 - 1) / (n - 1) if (n > n0 and n0 > 1) else g0
        w = (span - (n - 1) * gap) / n
    for i, k in enumerate(sorted(items)):
        el = items[k]
        bx, by, bw, bh = geom.get_box(el)
        x = span_l + i * (w + gap) if n > 1 else span_l + (span - w) / 2
        geom.set_group_box(el, x, by, w, bh * scale if proportional else bh)
    _reroute_connectors(items, conns, scale)
    return {"scale": scale}


def _place_keep_size(tree, units, conns, boxes0, n, n0, axis, slide_size, res):
    """Вертикальные ряды и «свободные» элементы: размер не меняется, меняется зазор (>= MIN_GAP_FACTOR)."""
    a = 0 if axis == "x" else 1
    s = a + 2
    bx = sorted(boxes0, key=lambda b: b[a])
    gaps = [bx[i + 1][a] - (bx[i][a] + bx[i][s]) for i in range(len(bx) - 1)]
    g0 = max(0, sum(gaps) / len(gaps)) if gaps else DEFAULT_GAP
    span_l = min(b[a] for b in boxes0)
    orig_r = max(b[a] + b[s] for b in boxes0)
    end = orig_r
    if slide_size:
        end = _span_limit(tree, [e for v in units.values() for e in v], a, slide_size, orig_r)
    order = sorted(units)
    ubox = {k: unit_box(units[k]) for k in order}
    total = sum(ubox[k][s] for k in order)
    gmin = g0 * MIN_GAP_FACTOR
    if n == 1:
        gap = 0.0
    elif n <= n0 or total + (n - 1) * g0 <= end - span_l:
        gap = g0
    else:
        gap = max(gmin, (end - span_l - total) / (n - 1))
    occupied = total + (n - 1) * gap
    if n > 1 and occupied > end - span_l + 1:
        res["warnings"].append(
            f"{n} элементов не помещаются в отведённое место даже с минимальным зазором "
            f"— разделите на слайды")
    pos = span_l
    if axis == "x" and n <= n0:
        pos = span_l + (orig_r - span_l - occupied) / 2
    for k in order:
        d = pos - ubox[k][a]
        _shift(units[k], d if a == 0 else 0, d if a == 1 else 0)
        pos += ubox[k][s] + gap
    if conns and axis == "y":
        _reroute_vertical(units, conns)
    elif conns:
        items = {k: v[0] for k, v in units.items() if len(v) == 1}
        if len(items) == len(units):
            _reroute_connectors(items, conns, 1.0)
    return {}


def _reroute_vertical(units, conns):
    for k in sorted(conns):
        if k + 1 not in units:
            continue
        a, b = unit_box(units[k]), unit_box(units[k + 1])
        x = a[0] + a[2] / 2
        y1, y2 = a[1] + a[3] + PAD, b[1] - PAD
        geom.set_box(conns[k], x, y1, 0, max(y2 - y1, 1))
        xf = geom.xfrm(conns[k])
        for attr in ("flipH", "flipV", "rot"):
            if attr in xf.attrib:
                del xf.attrib[attr]


def _anchor_box(item):
    """Опорная фигура коннектора: ребёнок '.num', иначе весь item."""
    if item.tag == geom.GRP:
        for el, ab, _ in geom.walk(list(item), _item_tf(item)):
            if (geom.name_of(el) or "").endswith(".num"):
                return ab
    return geom.get_box(item)


def _item_tf(item):
    # абсолютная система координат для детей группы
    b = geom.get_box(item)
    xf = geom.xfrm(item)
    co, ce = xf.find(qn("a:chOff")), xf.find(qn("a:chExt"))
    sx = b[2] / int(ce.get("cx")) if int(ce.get("cx")) else 1.0
    sy = b[3] / int(ce.get("cy")) if int(ce.get("cy")) else 1.0
    return (sx, sy, b[0] - int(co.get("x")) * sx, b[1] - int(co.get("y")) * sy)


def _reroute_connectors(items, conns, scale):
    ks = sorted(items)
    for k in sorted(conns):
        if k + 1 not in items:
            continue
        a, b = _anchor_box(items[k]), _anchor_box(items[k + 1])
        pad = PAD * scale
        x1, x2 = a[0] + a[2] + pad, b[0] - pad
        y = a[1] + a[3] / 2
        geom.set_box(conns[k], x1, y, max(x2 - x1, 1), 0)
        xf = geom.xfrm(conns[k])
        for attr in ("flipH", "flipV", "rot"):
            if attr in xf.attrib:
                del xf.attrib[attr]
