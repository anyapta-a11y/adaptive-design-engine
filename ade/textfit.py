"""Метрики текста по шрифтовым файлам (Pillow): перенос строк, высота, вместимость."""
from __future__ import annotations

import functools
import math
from dataclasses import dataclass
from pathlib import Path

from PIL import ImageFont
from pptx.oxml.ns import qn

LINE_FACTOR = 1.21     # межстрочный коэффициент Inter
REF = 200              # опорный размер загрузки шрифта
DEFAULT_PT = 18.0
SAMPLE = "Пример текста для оценки средней ширины символов в строке"
_DIRS = [Path(__file__).resolve().parent.parent / "fonts",
         Path("/usr/share/fonts/opentype/inter"), Path("/usr/share/fonts/truetype/inter"),
         Path("/usr/share/fonts/truetype/dejavu")]


@functools.lru_cache(maxsize=None)
def _font(name: str, bold: bool):
    key = (name or "Inter").replace(" ", "")
    stems = [f"{key}-{'Bold' if bold else 'Regular'}", key]
    for stem in stems + (["DejaVuSans-Bold"] if bold else ["DejaVuSans"]):
        for d in _DIRS:
            for ext in (".otf", ".ttf"):
                p = d / (stem + ext)
                if p.exists():
                    return ImageFont.truetype(str(p), REF)
    raise RuntimeError(f"Не найден шрифт {name!r} (ни в fonts/, ни в системных каталогах)")


def width(text: str, font: str, bold: bool, pt: float) -> float:
    return _font(font, bold).getlength(text) * pt / REF


@dataclass
class Para:
    text: str
    pt: float
    bold: bool
    font: str
    lnspc: float = 1.0
    before: float = 0.0
    after: float = 0.0


def txbody(el):
    """txBody фигуры (p:txBody) или ячейки таблицы (a:txBody)."""
    t = el.find(qn("p:txBody"))
    return t if t is not None else el.find(qn("a:txBody"))


def paras(txb, override_pt: float | None = None) -> list[Para]:
    out = []
    for p in txb.findall(qn("a:p")):
        runs = p.findall(qn("a:r"))
        text = "".join((r.findtext(qn("a:t")) or "") for r in runs)
        sizes, bold, font = [], False, None
        for r in runs:
            rpr = r.find(qn("a:rPr"))
            if rpr is None:
                continue
            if rpr.get("sz"):
                sizes.append(int(rpr.get("sz")) / 100)
            bold = bold or rpr.get("b") == "1"
            lat = rpr.find(qn("a:latin"))
            if lat is not None and font is None:
                font = lat.get("typeface")
        if not sizes:
            end = p.find(qn("a:endParaRPr"))
            if end is not None and end.get("sz"):
                sizes.append(int(end.get("sz")) / 100)
        pt = override_pt if override_pt else (max(sizes) if sizes else DEFAULT_PT)
        ppr = p.find(qn("a:pPr"))
        ln, bef, aft = 1.0, 0.0, 0.0
        if ppr is not None:
            e = ppr.find(qn("a:lnSpc") + "/" + qn("a:spcPct"))
            if e is not None:
                ln = int(e.get("val")) / 100000
            e = ppr.find(qn("a:spcBef") + "/" + qn("a:spcPts"))
            if e is not None:
                bef = int(e.get("val")) / 100
            e = ppr.find(qn("a:spcAft") + "/" + qn("a:spcPts"))
            if e is not None:
                aft = int(e.get("val")) / 100
        out.append(Para(text, pt, bold, font or "Inter", ln, bef, aft))
    return out


def current_pt(txb) -> float:
    ps = [p.pt for p in paras(txb)]
    return min(ps) if ps else DEFAULT_PT


def wrap(text: str, avail: float, font: str, bold: bool, pt: float):
    """(число строк, ширина самого длинного слова) при жадном переносе по пробелам."""
    words = text.split()
    if not words:
        return 1, 0.0
    longest = max(width(w, font, bold, pt) for w in words)
    lines, cur = 1, ""
    for w in words:
        cand = (cur + " " + w) if cur else w
        if width(cand, font, bold, pt) <= avail or not cur:
            cur = cand
        else:
            lines += 1
            cur = w
    # слово шире строки ломается на несколько строк
    extra = sum(max(0, math.ceil(width(w, font, bold, pt) / avail) - 1)
                for w in words) if avail > 0 else 0
    return lines + extra, longest


def measure(txb, avail_w: float, override_pt: float | None = None):
    """(высота текста pt, ширина длиннейшего слова pt) для области шириной avail_w pt."""
    h, longest = 0.0, 0.0
    for p in paras(txb, override_pt):
        lines, lw = wrap(p.text, avail_w, p.font, p.bold, p.pt)
        h += lines * p.pt * LINE_FACTOR * p.lnspc + p.before + p.after
        longest = max(longest, lw)
    return h, longest


def insets(txb, cell: bool = False, tcpr=None):
    """(left+right, top+bottom) в pt."""
    if cell:
        g = (lambda k, d: int(tcpr.get(k, d)) if tcpr is not None else d)
        return (g("marL", 91440) + g("marR", 91440)) / EMU, (g("marT", 45720) + g("marB", 45720)) / EMU
    bp = txb.find(qn("a:bodyPr"))
    g = (lambda k, d: int(bp.get(k, d)) if bp is not None else d)
    return (g("lIns", 91440) + g("rIns", 91440)) / EMU, (g("tIns", 45720) + g("bIns", 45720)) / EMU


EMU = 12700.0


def capacity(w_emu: float, h_emu: float, txb, scale: float = 1.0) -> int:
    """Оценка вместимости слота в символах (русский текст, перенос по словам)."""
    ps = paras(txb)
    if not ps:
        return 0
    p = ps[0]
    pt = p.pt * scale
    iw, ih = insets(txb)
    aw, ah = w_emu / EMU - iw, h_emu / EMU - ih
    per_char = width(SAMPLE, p.font, p.bold, pt) / len(SAMPLE)
    lines = max(0, int(ah // (pt * LINE_FACTOR * p.lnspc)))
    return int(lines * (aw / per_char) * 0.92)
