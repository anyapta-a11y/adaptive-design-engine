"""Поиск именованных слотов, диаграмм и статичного текста на любой глубине вложенности групп
(кроме внутренностей групп-элементов item[k])."""
from __future__ import annotations

import re

from pptx.oxml.ns import qn

from . import geom
from . import textfit as tf
from .ops.repeat import ITEM_RE

SLOT_RE = re.compile(r"^[a-z][a-z0-9_]*$")
CHART_RE = re.compile(r"^chart(?:_(\d+))?$")
LEGACY = ("title", "subtitle", "body", "note")


def is_slot_name(name) -> bool:
    return bool(name) and bool(SLOT_RE.match(name)) and name != "table" and not CHART_RE.match(name)


def _in_item_group(el) -> bool:
    p = el.getparent()
    while p is not None:
        if p.tag == geom.GRP and ITEM_RE.match(geom.name_of(p) or ""):
            return True
        p = p.getparent()
    return False


def scan(tree):
    """[(элемент, абсолютный бокс)] всех фигур вне групп item[k] (с вложенностью)."""
    top = [e for e in tree if e.tag in geom.SHAPE_TAGS]
    return [(e, ab) for e, ab, _ in geom.walk(top) if not _in_item_group(e)]


def text_of(txb) -> str:
    return "\n".join("".join((r.findtext(qn("a:t")) or "") for r in p.findall(qn("a:r")))
                     for p in txb.findall(qn("a:p")))


def slot_elements(tree) -> list:
    """[(имя, элемент, бокс)] текстовых слотов; или легаси-имя без текста (для ошибки)."""
    out = []
    for e, ab in scan(tree):
        n = geom.name_of(e)
        if e.tag == geom.SP and is_slot_name(n) and (tf.txbody(e) is not None
                                                      and (n in LEGACY or text_of(tf.txbody(e)).strip())):
            out.append((n, e, ab))
    return out


def chart_elements(tree) -> dict:
    """{имя: graphicFrame} для chart, chart_2, ..."""
    return {geom.name_of(e): e for e, _ in scan(tree)
            if e.tag == geom.GF and CHART_RE.match(geom.name_of(e) or "")
            and e.find(".//" + qn("c:chart")) is not None}


def chart_sort_key(name: str) -> int:
    m = CHART_RE.match(name)
    return int(m.group(1)) if m.group(1) else 1


def static_text(tree) -> list:
    slot_names = {n for n, _, _ in slot_elements(tree)}
    out = []
    for e, _ in scan(tree):
        n = geom.name_of(e) or ""
        if e.tag != geom.SP or n in slot_names or n.startswith("item[") or n == "table" or CHART_RE.match(n):
            continue
        txb = tf.txbody(e)
        if txb is None:
            continue
        t = " ".join(text_of(txb).split())
        if t:
            out.append({"shape": n, "text": t[:40]})
    return out
