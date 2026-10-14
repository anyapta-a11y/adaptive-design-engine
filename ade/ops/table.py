"""Нативная таблица: ±строки/столбцы клонированием a:tr/a:tc, сохранение чередования, высота."""
from __future__ import annotations

import copy

from pptx.oxml.ns import qn

from .. import textfit as tf
from ..errors import PlanError
from .text import set_text

BOTTOM_MARGIN = 365760      # 0.4"


def tbl_of(gf):
    return gf.find(".//" + qn("a:tbl"))


def col_widths(tbl):
    return [int(g.get("w")) for g in tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol"))]


def row_needs(tbl) -> list[int]:
    """Минимальная высота каждой строки (EMU), чтобы влез текст ячеек."""
    widths = col_widths(tbl)
    needs = []
    for tr in tbl.findall(qn("a:tr")):
        need = 0.0
        for i, tc in enumerate(tr.findall(qn("a:tc"))):
            txb = tc.find(qn("a:txBody"))
            tcpr = tc.find(qn("a:tcPr"))
            iw, ih = tf.insets(txb, cell=True, tcpr=tcpr)
            h, _ = tf.measure(txb, widths[i] / tf.EMU - iw)
            need = max(need, h + ih)
        needs.append(int(need * tf.EMU) + 1)
    return needs


def estimated_height(tbl) -> int:
    trs = tbl.findall(qn("a:tr"))
    return sum(max(int(tr.get("h")), n) for tr, n in zip(trs, row_needs(tbl)))


def _adjust_cols(tbl, cols: int) -> None:
    grid = tbl.find(qn("a:tblGrid"))
    gcs = grid.findall(qn("a:gridCol"))
    total = sum(int(g.get("w")) for g in gcs)
    while len(gcs) < cols:
        grid.append(copy.deepcopy(gcs[-1])); gcs = grid.findall(qn("a:gridCol"))
        for tr in tbl.findall(qn("a:tr")):
            tcs = tr.findall(qn("a:tc")); tcs[-1].addnext(copy.deepcopy(tcs[-1]))
    while len(gcs) > cols:
        grid.remove(gcs[-1]); gcs = grid.findall(qn("a:gridCol"))
        for tr in tbl.findall(qn("a:tr")):
            tr.remove(tr.findall(qn("a:tc"))[-1])
    cur = sum(int(g.get("w")) for g in gcs)
    acc = 0
    for i, g in enumerate(gcs):
        w = int(int(g.get("w")) * total / cur) if i < len(gcs) - 1 else total - acc
        g.set("w", str(w)); acc += w


def fill_table(gf, rows: list[list[str]], slide_h: int) -> dict:
    """gf: p:graphicFrame с таблицей. rows[0] — шапка. Возвращает {'warnings': [...]}."""
    if not rows or any(len(r) != len(rows[0]) for r in rows):
        raise PlanError("Таблица плана: все строки должны иметь одинаковое число ячеек")
    if len(rows) < 2:
        raise PlanError("Таблица плана: нужна шапка и хотя бы одна строка данных")
    tbl = tbl_of(gf)
    trs = tbl.findall(qn("a:tr"))
    if len(trs) < 2:
        raise PlanError("В образце библиотеки таблица должна содержать шапку и строку данных")
    _adjust_cols(tbl, len(rows[0]))
    while len(trs) < len(rows):
        src = trs[-2] if len(trs) >= 3 else trs[-1]    # через одну: сохраняет чередование заливки
        trs[-1].addnext(copy.deepcopy(src)); trs = tbl.findall(qn("a:tr"))
    while len(trs) > len(rows):
        tbl.remove(trs[-1]); trs = tbl.findall(qn("a:tr"))
    for tr, row in zip(trs, rows):
        for tc, val in zip(tr.findall(qn("a:tc")), row):
            set_text(tc.find(qn("a:txBody")), str(val))
    # высота: исходная, если влезает; иначе сжатие строк до минимума по тексту; дальше — предупреждение
    top = int(gf.find(qn("p:xfrm")).find(qn("a:off")).get("y"))
    avail = slide_h - BOTTOM_MARGIN - top
    needs = row_needs(tbl)
    desired = [max(int(tr.get("h")), n) for tr, n in zip(trs, needs)]
    warnings = []
    if sum(desired) > avail:
        smin = sum(needs)
        if smin >= avail:
            desired = needs
            warnings.append("таблица не помещается по высоте даже с минимальными строками")
        else:
            k = (avail - smin) / (sum(desired) - smin)
            desired = [int(n + (d - n) * k) for d, n in zip(desired, needs)]
    for tr, h in zip(trs, desired):
        tr.set("h", str(h))
    gf.find(qn("p:xfrm")).find(qn("a:ext")).set("cy", str(sum(desired)))
    return {"warnings": warnings}
