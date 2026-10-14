"""Нарисованные столбцы/полосы item[n].bar: длина пропорциональна значению из плана."""
from __future__ import annotations

import re

from .. import geom
from ..errors import PlanError

BAR_RE = re.compile(r"^item\[(\d+)\]\.bar$")
TOL = 25400     # 2pt: допуск «элемент целиком за свободным краем столбца»


def dim_for_axis(axis: str) -> str:
    """Столбцы (ось повторения x) растут по высоте, полосы (ось y) — по ширине."""
    return "h" if axis == "x" else "w"


def find_bars(tree) -> dict:
    out = {}
    for c in tree.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"):
        m = BAR_RE.match(c.get("name") or "")
        if m:
            out[int(m.group(1))] = c.getparent().getparent()
    return out


def _len(el, dim):
    b = geom.get_box(el)
    return b[3] if dim == "h" else b[2]


def reference_length(tree, dim: str) -> int:
    """Эталон «полного» столбца = самый длинный столбец макета (до клонирования)."""
    bars = find_bars(tree)
    return max((_len(e, dim) for e in bars.values()), default=0)


def scale_bars(tree, values: dict, ref: int, dim: str, bar_max=None) -> dict:
    """values: {номер item: число}. Нижний (или левый) край фиксирован; сопутствующие фигуры,
    лежащие целиком за свободным краем (значение над столбцом), смещаются вместе с ним."""
    warns = []
    for i, v in values.items():
        if isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0:
            raise PlanError(f"items[{i}].bar: нужно неотрицательное число, получено {v!r}")
    explicit = bar_max is not None
    mx = float(bar_max) if explicit else max(values.values(), default=0)
    bars = find_bars(tree)
    for i, v in values.items():
        el = bars[i]
        x, y, w, h = geom.get_box(el)
        if explicit and v > mx:
            warns.append(f"items[{i}].bar={v:g} больше bar_max={mx:g}: столбец выше эталона")
        k = (v / mx) if mx > 0 else 0.0
        if dim == "h":
            nh = ref * k
            geom.set_box(el, x, y + h - nh, w, nh)
            edge_old, delta = y, (y + h - nh) - y
        else:
            nw = ref * k
            geom.set_box(el, x, y, nw, h)
            edge_old, delta = x + w, nw - w
        name = f"item[{i}]."
        for sib in el.getparent():
            if sib is el or sib.tag not in geom.SHAPE_TAGS or not (geom.name_of(sib) or "").startswith(name):
                continue
            b = geom.get_box(sib)
            if not b:
                continue
            if dim == "h" and b[1] + b[3] <= edge_old + TOL:
                geom.set_group_box(sib, b[0], b[1] + delta, b[2], b[3])
            elif dim == "w" and b[0] >= edge_old - TOL:
                geom.set_group_box(sib, b[0] + delta, b[1], b[2], b[3])
    return {"max": mx, "ref": ref, "dim": dim, "warnings": warns}
