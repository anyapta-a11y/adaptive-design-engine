"""Заполнение текста с сохранением форматирования и единая подгонка кегля."""
from __future__ import annotations

import copy

from pptx.oxml.ns import qn

from .. import textfit as tf

MIN_SCALE = 0.75


def _is_bold(rpr) -> bool:
    if rpr is None:
        return False
    if rpr.get("b") in ("1", "true"):
        return True
    lat = rpr.find(qn("a:latin"))
    return lat is not None and "bold" in (lat.get("typeface") or "").lower()


def _segments(line: str) -> list:
    """'a **b** c' -> [('a ', False), ('b', True), (' c', False)]; маркеры не попадают в результат."""
    out, bold = [], False
    for i, part in enumerate(line.split("**")):
        if i:
            bold = not bold
        if part:
            out.append((part, bold))
    return out


def set_text(txb, text: str) -> None:
    """Текст в первый run (форматирование сохраняется); '\\n' -> новые абзацы с тем же форматом.
    '**...**' -> выделение: rPr жирного run шаблона (иначе rPr обычного с b=1)."""
    ps = txb.findall(qn("a:p"))
    tmpl = copy.deepcopy(ps[0])
    runs = tmpl.findall(qn("a:r"))
    end = tmpl.find(qn("a:endParaRPr"))
    rprs = [r.find(qn("a:rPr")) for r in runs]
    if not runs:
        base = copy.deepcopy(end) if end is not None else None
        if base is not None:
            base.tag = qn("a:rPr")
        normal, bold_src = base, None
    else:
        bold_src = next((x for x in rprs if _is_bold(x)), None)
        normal = next((x for x in rprs if not _is_bold(x) and x is not None), None)
        if normal is None:
            normal = next((x for x in rprs if not _is_bold(x)), None)    # run без rPr
        if normal is None:
            normal = rprs[0]                                              # всё жирное
    for extra in tmpl.findall(qn("a:r")) + tmpl.findall(qn("a:br")) + tmpl.findall(qn("a:fld")):
        tmpl.remove(extra)
    # rPr для первого run шаблона, как раньше (без маркеров)
    plain = rprs[0] if runs else normal
    for p in ps:
        txb.remove(p)
    for line in text.split("\n"):
        p = copy.deepcopy(tmpl)
        segs = _segments(line)
        has_marks = "**" in line
        if not segs:
            segs = [("", False)]
        for t, b in segs:
            r = p.makeelement(qn("a:r"), {})
            if has_marks:
                src = (bold_src if bold_src is not None else normal) if b else normal
                rpr = copy.deepcopy(src) if src is not None else None
                if b and bold_src is None:
                    if rpr is None:
                        rpr = r.makeelement(qn("a:rPr"), {})
                    rpr.set("b", "1")
            else:
                rpr = copy.deepcopy(plain) if plain is not None else None
            if rpr is not None:
                r.append(rpr)
            te = r.makeelement(qn("a:t"), {}); te.text = t; r.append(te)
            (p.find(qn("a:endParaRPr")).addprevious(r) if p.find(qn("a:endParaRPr")) is not None
             else p.append(r))
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
