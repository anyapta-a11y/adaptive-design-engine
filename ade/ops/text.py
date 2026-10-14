"""Заполнение текста с сохранением форматирования и единая подгонка кегля."""
from __future__ import annotations

import copy

from pptx.oxml.ns import qn

from .. import textfit as tf

MIN_SCALE = 0.75


def set_text(txb, text: str) -> None:
    """Текст в первый run (форматирование сохраняется); '\\n' -> новые абзацы с тем же форматом."""
    ps = txb.findall(qn("a:p"))
    tmpl = copy.deepcopy(ps[0])
    runs = tmpl.findall(qn("a:r"))
    for extra in tmpl.findall(qn("a:br")) + tmpl.findall(qn("a:fld")) + runs[1:]:
        tmpl.remove(extra)
    if not runs:
        r = tmpl.makeelement(qn("a:r"), {})
        end = tmpl.find(qn("a:endParaRPr"))
        if end is not None:
            rpr = copy.deepcopy(end); rpr.tag = qn("a:rPr"); r.append(rpr)
        t = r.makeelement(qn("a:t"), {}); r.append(t)
        (end.addprevious(r) if end is not None else tmpl.append(r))
    for p in ps:
        txb.remove(p)
    for line in text.split("\n"):
        p = copy.deepcopy(tmpl)
        p.find(qn("a:r")).find(qn("a:t")).text = line
        txb.append(p)


def set_font_size(txb, pt: float) -> None:
    sz = str(int(round(pt * 100)))
    for p in txb.findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            if rpr is None:
                rpr = r.makeelement(qn("a:rPr"), {}); r.insert(0, rpr)
            rpr.set("sz", sz)
        end = p.find(qn("a:endParaRPr"))
        if end is not None:
            end.set("sz", sz)


def read_text(txb) -> str:
    return "\n".join("".join((r.findtext(qn("a:t")) or "") for r in p.findall(qn("a:r")))
                     for p in txb.findall(qn("a:p")))


def fits(txb, w_emu, h_emu, pt=None) -> tuple[bool, bool]:
    """(по высоте ок, по ширине слов ок) для текстового тела."""
    iw, ih = tf.insets(txb)
    aw, ah = w_emu / tf.EMU - iw, h_emu / tf.EMU - ih
    h, longest = tf.measure(txb, aw, pt)
    return h <= ah + 0.5, longest <= aw + 0.5


def fit_uniform(members, floor: float = MIN_SCALE) -> dict:
    """members: [(txb, w_emu, h_emu)]. Единый кегль для группы: шаг 1pt вниз, не ниже 75%.
    Текст не сокращается. Возвращает {from, to, fits}."""
    members = [m for m in members if read_text(m[0]).strip()]
    if not members:
        return {"from": None, "to": None, "fits": True}
    orig = min(tf.current_pt(m[0]) for m in members)
    lo = orig * floor

    def ok(pt):
        return all(all(fits(t, w, h, pt)) for t, w, h in members)

    pt = orig
    while not ok(pt) and pt - 1 >= lo - 1e-9:
        pt -= 1
    if not ok(pt) and pt > lo:
        pt = lo
    good = ok(pt)
    if pt != orig:
        for t, _, _ in members:
            set_font_size(t, pt)
    return {"from": orig, "to": pt, "fits": good}
