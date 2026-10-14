"""Нативная диаграмма: replace_data (кэш + встроенный xlsx), number_format, цвета серий."""
from __future__ import annotations

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor

from ..errors import PlanError

EXTRA_PALETTE = ["7A8CA5", "C9A227", "4F9D8A", "B5523B", "6B5B95"]


def validate(spec: dict) -> None:
    cats, series = spec.get("categories"), spec.get("series")
    if not cats or not series:
        raise PlanError("chart: нужны categories и series")
    for s in series:
        if len(s.get("values", [])) != len(cats):
            raise PlanError(f"chart: в серии {s.get('name')!r} число значений != числу категорий")


def _rgb(ser):
    try:
        return str(ser.format.fill.fore_color.rgb)
    except Exception:
        return None


def fill_chart(chart, spec: dict) -> None:
    """chart: pptx.chart.chart.Chart (объект нового слайда)."""
    validate(spec)
    plot = chart.plots[0]
    old = [_rgb(s) for s in plot.series]
    fmt = spec.get("number_format")
    if fmt is None:
        try:
            fmt = plot.data_labels.number_format if plot.has_data_labels else "General"
        except Exception:
            fmt = "General"
    cd = CategoryChartData(number_format=fmt)
    cd.categories = [str(c) for c in spec["categories"]]
    for s in spec["series"]:
        cd.add_series(str(s["name"]), list(s["values"]))
    chart.replace_data(cd)
    plot = chart.plots[0]
    given = spec.get("colors") or []
    known = [c for c in old if c]
    for i, ser in enumerate(plot.series):
        color = given[i] if i < len(given) else (old[i] if i < len(old) and old[i] else None)
        if color is None:
            color = EXTRA_PALETTE[(i - len(known)) % len(EXTRA_PALETTE)] if known else EXTRA_PALETTE[i % len(EXTRA_PALETTE)]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#").upper())
    try:
        plot.vary_by_categories = False
        if plot.has_data_labels:
            plot.data_labels.number_format = fmt
            plot.data_labels.number_format_is_linked = False
        chart.value_axis.tick_labels.number_format = fmt
        chart.value_axis.tick_labels.number_format_is_linked = False
    except Exception:
        pass    # диаграммы без осей (pie)
