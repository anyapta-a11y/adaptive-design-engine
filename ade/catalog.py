"""Разметка library.pptx -> catalog.json (с оценкой вместимости слотов)."""
from __future__ import annotations

import json
import re

from pptx import Presentation
from pptx.oxml.ns import qn

from . import geom
from . import textfit as tf
from .errors import MarkupError
from .ops.repeat import CONN_RE, ITEM_RE
from .ops import table as tbl_ops

FAMILIES = {"title", "cards", "process", "kpi", "table", "chart", "text"}
SLOTS = ("title", "subtitle", "body", "note")
FIELD_RE = re.compile(r"^item\[(\d+)\]\.(\w+)$")
NEEDS = {"cards": "items", "process": "items", "kpi": "items", "table": "table",
         "chart": "chart", "text": "body", "title": "title"}


def parse_note(text: str) -> dict:
    first = (text or "").strip().splitlines()[0].strip() if (text or "").strip() else ""
    if not first.lower().startswith("ade:"):
        return {}
    out = {}
    for part in first[4:].split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def _cap(ab, el):
    txb = tf.txbody(el)
    if txb is None:
        return None
    p = tf.paras(txb)
    return {"capacity_chars": tf.capacity(ab[2], ab[3], txb),
            "capacity_chars_min_font": tf.capacity(ab[2], ab[3], txb, 0.75),
            "font_pt": p[0].pt if p else tf.DEFAULT_PT}


def analyse_slide(slide, num: int, slide_h: int) -> dict:
    meta = {}
    if slide.has_notes_slide:
        meta = parse_note(slide.notes_slide.notes_text_frame.text)
    where = f"слайд {num}"
    sid = meta.get("id")
    if not meta:
        raise MarkupError(f"{where}: нет заметки вида 'ade: id=…; family=…; purpose=…'")
    for key in ("id", "family", "purpose"):
        if not meta.get(key):
            raise MarkupError(f"{where}: в заметке 'ade:' отсутствует {key}")
    where = f"слайд {num} ({sid})"
    fam = meta["family"]
    if fam not in FAMILIES:
        raise MarkupError(f"{where}: неизвестное семейство {fam!r}; допустимо: {sorted(FAMILIES)}")

    tree = slide.shapes._spTree
    top = [el for el in tree if el.tag in geom.SHAPE_TAGS]
    ab_of = {id(e): ab for e, ab, _ in geom.walk(top)}
    seen: dict[str, object] = {}
    for el in top:
        n = geom.name_of(el)
        if n in seen:
            raise MarkupError(f"{where}: имя фигуры {n!r} встречается дважды")
        seen[n] = el
        if n.startswith("item[") and not ITEM_RE.match(n):
            raise MarkupError(f"{where}: некорректное имя {n!r}, ожидается item[N]")
        if n.startswith("connector[") and not CONN_RE.match(n):
            raise MarkupError(f"{where}: некорректное имя {n!r}, ожидается connector[N]")

    slots = []
    for s in SLOTS:
        if s in seen:
            if tf.txbody(seen[s]) is None:
                raise MarkupError(f"{where}: {s!r} не содержит текста")
            slots.append({"name": s, **_cap(ab_of[id(seen[s])], seen[s])})

    nums = sorted(int(ITEM_RE.match(n).group(1)) for n in seen if ITEM_RE.match(n))
    if nums != list(range(1, len(nums) + 1)):
        raise MarkupError(f"{where}: нумерация item должна идти подряд с 1, найдено {nums}")
    fields_by_item: dict[int, dict] = {}
    for k in nums:
        el = seen[f"item[{k}]"]
        fields: dict[str, dict] = {}
        for c, ab, _ in geom.walk(list(el), _item_tf(el)) if el.tag == geom.GRP else []:
            cn = geom.name_of(c) or ""
            m = FIELD_RE.match(cn)
            if cn.startswith("item["):
                if not m or int(m.group(1)) != k:
                    raise MarkupError(f"{where}: поле {cn!r} внутри item[{k}] названо неверно")
                if tf.txbody(c) is None:
                    raise MarkupError(f"{where}: поле {cn!r} не содержит текста")
                fields[m.group(2)] = _cap(ab, c)
        fields_by_item[k] = fields
    if nums:
        ref = set(fields_by_item[1])
        for k, f in fields_by_item.items():
            if set(f) != ref:
                raise MarkupError(f"{where}: поля item[{k}] {sorted(f)} не совпадают с item[1] {sorted(ref)}")
    conns = sorted(int(CONN_RE.match(n).group(1)) for n in seen if CONN_RE.match(n))
    if conns and (conns != list(range(1, len(conns) + 1)) or len(conns) != max(len(nums) - 1, 0)):
        raise MarkupError(f"{where}: connector должны быть номерами 1..{len(nums) - 1} "
                          f"(по числу промежутков между item), найдено {conns}")

    table = chart = None
    if "table" in seen:
        tb = tbl_ops.tbl_of(seen["table"]) if seen["table"].tag == geom.GF else None
        if tb is None:
            raise MarkupError(f"{where}: фигура 'table' не является нативной таблицей")
        trs = tb.findall(qn("a:tr"))
        cell = trs[min(1, len(trs) - 1)].find(qn("a:tc"))
        p = tf.paras(cell.find(qn("a:txBody")))
        pt = p[0].pt if p else tf.DEFAULT_PT
        min_h = (pt * tf.LINE_FACTOR + 7.2) * tf.EMU
        top_y = geom.get_box(seen["table"])[1]
        table = {"cols": len(tb.find(qn("a:tblGrid")).findall(qn("a:gridCol"))), "rows": len(trs),
                 "max_rows": int((slide_h - tbl_ops.BOTTOM_MARGIN - top_y) // min_h)}
    if "chart" in seen:
        el = seen["chart"]
        if el.tag != geom.GF or el.find(".//" + qn("c:chart")) is None:
            raise MarkupError(f"{where}: фигура 'chart' не является нативной диаграммой")
        ch = next(s for s in slide.shapes if s.name == "chart").chart
        chart = {"type": str(ch.chart_type).split(".")[-1].split(" ")[0],
                 "series": len(list(ch.plots[0].series)),
                 "categories": len(list(ch.plots[0].categories))}
    need = NEEDS[fam]
    ok = {"items": bool(nums), "table": table is not None, "chart": chart is not None,
          "body": "body" in seen, "title": "title" in seen}[need]
    if not ok:
        raise MarkupError(f"{where}: для семейства {fam!r} нужна фигура/слоты {need!r}")
    return {"id": sid, "family": fam, "purpose": meta["purpose"], "slide_index": num - 1,
            "slots": slots, "item_count": len(nums),
            "item_fields": [{"name": n, **v} for n, v in fields_by_item.get(1, {}).items()],
            "table": table, "chart": chart}


def _item_tf(el):
    from .ops.repeat import _item_tf as f
    return f(el)


def build_catalog(path: str) -> dict:
    prs = Presentation(path)
    entries, ids = [], set()
    for i, s in enumerate(prs.slides, 1):
        e = analyse_slide(s, i, prs.slide_height)
        if e["id"] in ids:
            raise MarkupError(f"Дублирующийся id {e['id']!r} (слайд {i})")
        ids.add(e["id"])
        entries.append(e)
    fams: dict[str, list] = {}
    for e in entries:
        fams.setdefault(e["family"], []).append(e["id"])
    return {"library": str(path), "families": fams, "slides": entries}


def dumps(cat: dict) -> str:
    return json.dumps(cat, ensure_ascii=False, indent=2)
