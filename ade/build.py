"""plan.json + library.pptx -> out.pptx (+ out.pptx.build.json со сведениями о сборке)."""
from __future__ import annotations

import json
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

from . import geom
from . import textfit as tf
from .catalog import SLOTS, build_catalog
from .errors import PlanError
from .ops import chart as chart_ops
from .ops import table as table_ops
from .ops.repeat import ITEM_RE, adapt_items
from .ops.slides import copy_slide, drop_slides
from .ops.text import fit_uniform, set_text

PROPORTIONAL_FAMILIES = {"process"}    # у этих семейств элементы уменьшаются пропорционально


def sidecar_path(out) -> Path:
    return Path(str(out) + ".build.json")


def select_variant(cat: dict, ps: dict, where: str = "слайд плана"):
    """-> (entry, action, n_items). action: exact | clone | trim | fixed."""
    entries = cat["slides"]
    n = len(ps["items"]) if "items" in ps else None
    if n is not None and n < 1:
        raise PlanError(f"{where}: items пуст")
    if ps.get("id"):
        e = next((x for x in entries if x["id"] == ps["id"]), None)
        if e is None:
            raise PlanError(f"{where}: в библиотеке нет слайда id={ps['id']!r}")
        pool = [e]
    elif ps.get("family"):
        pool = sorted((x for x in entries if x["family"] == ps["family"]), key=lambda x: x["item_count"])
        if not pool:
            raise PlanError(f"{where}: в библиотеке нет семейства {ps['family']!r}")
    else:
        raise PlanError(f"{where}: нужен id или family")
    if n is None:
        e = pool[0]
        if e["item_count"]:
            raise PlanError(f"{where}: для {e['id']!r} нужен список items")
        return e, "fixed", None
    pool = [x for x in pool if x["item_count"] > 0]
    if not pool:
        raise PlanError(f"{where}: слайд не поддерживает items")
    exact = [x for x in pool if x["item_count"] == n]
    if exact:
        return exact[0], "exact", n
    smaller = [x for x in pool if x["item_count"] < n]
    if smaller:
        return max(smaller, key=lambda x: x["item_count"]), "clone", n
    return min(pool, key=lambda x: x["item_count"]), "trim", n


def _top(tree):
    return {geom.name_of(e): e for e in tree if e.tag in geom.SHAPE_TAGS}


def _find(root, name):
    for c in root.iter(qn("p:cNvPr")):
        if c.get("name") == name:
            return c.getparent().getparent()
    return None


def _check_text(v, where):
    if not isinstance(v, str):
        raise PlanError(f"{where}: текст должен быть строкой, получено {type(v).__name__}")
    return v


def build_slide(prs, src, entry, ps, action, num):
    where = f"слайд плана {num}"
    new = copy_slide(prs, src)
    tree = new.shapes._spTree
    info = {"variant": entry["id"], "action": action, "cloned": 0, "deleted": 0, "fonts": [], "warnings": []}
    if action in ("clone", "trim", "exact") and ps.get("items") is not None:
        r = adapt_items(tree, len(ps["items"]), entry["family"] in PROPORTIONAL_FAMILIES)
        info["cloned"], info["deleted"] = r["cloned"], r["deleted"]
        if r["cloned"] or r["deleted"]:
            info["action"] = "clone" if r["cloned"] else "trim"
    top = _top(tree)
    # поля верхнего уровня
    fields = ps.get("fields", {})
    for k in fields:
        if k not in SLOTS or k not in top:
            raise PlanError(f"{where}: слота {k!r} нет в слайде {entry['id']!r}")
    for k in SLOTS:
        if k in top:
            if k in fields:
                set_text(tf.txbody(top[k]), _check_text(fields[k], f"{where}.fields.{k}"))
            else:
                tree.remove(top[k])
    # элементы
    items = ps.get("items") or []
    allowed = {f["name"] for f in entry["item_fields"]}
    for i, it in enumerate(items, 1):
        extra = set(it) - allowed
        if extra:
            raise PlanError(f"{where}: item {i}: поля {sorted(extra)} нет в {entry['id']!r} "
                            f"(есть: {sorted(allowed)})")
        for f in sorted(allowed):
            el = _find(tree, f"item[{i}].{f}")
            if f in it:
                set_text(tf.txbody(el), _check_text(it[f], f"{where}.items[{i}].{f}"))
            elif f == "num":
                set_text(tf.txbody(el), str(i))
            else:
                el.getparent().remove(el)
    # таблица / диаграмма
    if entry["table"]:
        if "table" not in ps:
            raise PlanError(f"{where}: для слайда {entry['id']!r} нужен ключ table")
        r = table_ops.fill_table(_top(tree)["table"], [[str(c) for c in row] for row in ps["table"]],
                                 prs.slide_height)
        info["warnings"] += r["warnings"]
    elif "table" in ps:
        raise PlanError(f"{where}: в слайде {entry['id']!r} нет таблицы")
    if entry["chart"]:
        if "chart" not in ps:
            raise PlanError(f"{where}: для слайда {entry['id']!r} нужен ключ chart")
        shape = next(s for s in new.shapes if s.name == "chart")
        chart_ops.fill_chart(shape.chart, ps["chart"])
    elif "chart" in ps:
        raise PlanError(f"{where}: в слайде {entry['id']!r} нет диаграммы")
    # подгонка текста: единый кегль для однотипных полей
    groups: dict[str, list] = {}
    for e, b, _ in geom.walk([e for e in tree if e.tag in geom.SHAPE_TAGS]):
        n = geom.name_of(e) or ""
        txb = tf.txbody(e) if e.tag in (geom.SP,) else None
        if txb is None:
            continue
        m = n.split("].", 1)
        key = f"item.{m[1]}" if ITEM_RE.match(m[0] + "]") and len(m) == 2 else n
        if n in SLOTS or key.startswith("item."):
            groups.setdefault(key, []).append((txb, b[2], b[3]))
    for key, members in groups.items():
        r = fit_uniform(members)
        if r["from"] is not None and r["to"] != r["from"]:
            info["fonts"].append({"slot": key, "from": r["from"], "to": r["to"]})
        if not r["fits"]:
            info["warnings"].append(f"текст не помещается в {key} даже при {r['to']:g}pt (75%); текст не сокращён")
    return new, info


def build(library, plan, out) -> dict:
    cat = build_catalog(str(library))
    if not isinstance(plan, dict):
        plan = json.loads(Path(plan).read_text(encoding="utf-8"))
    if not plan.get("slides"):
        raise PlanError("В плане нет слайдов (ключ 'slides')")
    prs = Presentation(str(library))
    lib_slides = list(prs.slides)
    report = []
    for num, ps in enumerate(plan["slides"], 1):
        entry, action, _ = select_variant(cat, ps, f"слайд плана {num}")
        _, info = build_slide(prs, lib_slides[entry["slide_index"]], entry, ps, action, num)
        report.append(info)
    drop_slides(prs, lib_slides)
    prs.save(str(out))
    data = {"library": str(library), "slides": report}
    sidecar_path(out).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data
