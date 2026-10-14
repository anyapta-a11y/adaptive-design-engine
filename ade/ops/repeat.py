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


def adapt_items(tree, target: int, proportional: bool = False) -> dict:
    """Приводит число item[n] к target. Возвращает {'cloned': n, 'deleted': n}."""
    items, conns = _indexed(tree, ITEM_RE), _indexed(tree, CONN_RE)
    n0 = len(items)
    if n0 == 0:
        raise AdeError("В слайде нет item[n]")
    if target < 1:
        raise AdeError("Число элементов должно быть >= 1")
    if target == n0:
        return {"cloned": 0, "deleted": 0}
    boxes = [geom.get_box(items[k]) for k in sorted(items)]
    span_l = min(b[0] for b in boxes)
    span_r = max(b[0] + b[2] for b in boxes)
    xs = sorted(boxes)
    gaps = [xs[i + 1][0] - (xs[i][0] + xs[i][2]) for i in range(len(xs) - 1)]
    g0 = max(0, sum(gaps) / len(gaps)) if gaps else DEFAULT_GAP
    w0 = boxes[0][2]
    cloned = deleted = 0
    while len(items) < target:
        _clone_last(tree, items, "item")
        if conns:
            _clone_last(tree, conns, "connector")
        cloned += 1
    while len(items) > target:
        k = max(items)
        tree.remove(items.pop(k))
        deleted += 1
        if conns:
            tree.remove(conns.pop(max(conns)))
    n = target
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
    return {"cloned": cloned, "deleted": deleted, "scale": scale}


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
