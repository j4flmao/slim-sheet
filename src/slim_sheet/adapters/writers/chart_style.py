"""Clean, brand-coloured chart styling for openpyxl charts (PLAN §16.4).

openpyxl's defaults fight a clean business chart: gap width 150, legend on the right, major
gridlines, rounded corners, and axes Excel 365 hides unless `delete` is written explicitly.
Every chart slim-sheet builds goes through `style_chart`, which overrides all of them.
"""

from __future__ import annotations

from typing import Any

from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
)

from slim_sheet.core.a1 import parse_range
from slim_sheet.core.design import CONTEXT_GREY, Brand
from slim_sheet.core.refs import split_sheet

LINE_EMU = 28575  # 2.25 pt
RULE_EMU = 9525  # 0.75 pt
ROUND_KINDS = ("pie", "doughnut")
MAX_LABELLED_POINTS = 12


def style_chart(
    chart: Any,
    kind: str,
    brand: Brand,
    title: str | None = None,
    highlight: int | None = None,
    value_format: str | None = None,
) -> None:
    """Strip chart junk and colour series from the brand; `highlight` greys all other bars."""
    chart.roundedCorners = False
    chart.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    chart.style = None
    if title:
        chart.title = _title(title, brand)
    n_series = len(chart.series)
    if kind in ROUND_KINDS:
        _style_round(chart, brand)
        return
    _style_axes(chart, kind, brand, value_format)
    chart.legend = Legend(legendPos="t") if n_series > 1 else None
    for index, series in enumerate(chart.series):
        _colour_series(series, kind, brand.palette[index % len(brand.palette)])
    if kind == "bar":  # horizontal bars: first category on top, value axis at the bottom
        chart.x_axis.scaling.orientation = "maxMin"
        chart.y_axis.crosses = "max"
    if kind in ("column", "bar"):
        chart.gapWidth = 60 if n_series == 1 else 80
        chart.overlap = -10 if n_series > 1 else None
        if n_series == 1:
            _single_series_bars(chart, brand, highlight, value_format)


def _style_axes(chart: Any, kind: str, brand: Brand, value_format: str | None) -> None:
    axis_line = GraphicalProperties(ln=LineProperties(solidFill=brand.rule, w=RULE_EMU))
    for axis in (chart.x_axis, chart.y_axis):
        axis.delete = False  # Excel 365 hides axes when <c:delete> is missing
        axis.txPr = _text(900, brand.muted, bold=False)
        axis.spPr = axis_line
        axis.majorGridlines = None
    if kind not in ("column", "bar"):  # bars carry data labels instead of gridlines
        chart.y_axis.majorGridlines = _gridlines(brand)
    if value_format:
        chart.y_axis.number_format = value_format


def _gridlines(brand: Brand) -> Any:
    return ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=brand.rule, w=RULE_EMU)))


def _colour_series(series: Any, kind: str, hex_colour: str) -> None:
    props = series.graphicalProperties
    if props is None:
        series.graphicalProperties = props = GraphicalProperties()
    if kind in ("line", "scatter"):
        props.line = LineProperties(solidFill=hex_colour, w=LINE_EMU)
        series.smooth = False
        if kind == "line":
            series.marker.symbol = "none"
        else:
            series.marker.graphicalProperties = GraphicalProperties(solidFill=hex_colour)
            props.line = LineProperties(noFill=True)
        return
    props.solidFill = hex_colour
    props.line = LineProperties(noFill=True)


def _single_series_bars(
    chart: Any, brand: Brand, highlight: int | None, value_format: str | None
) -> None:
    series = chart.series[0]
    points = _point_count(series)
    if highlight is not None and 0 <= highlight < points:
        series.graphicalProperties.solidFill = CONTEXT_GREY
        point = DataPoint(idx=highlight)
        point.graphicalProperties = GraphicalProperties(solidFill=brand.primary)
        series.dPt.append(point)
    if 0 < points <= MAX_LABELLED_POINTS:
        labels = DataLabelList()
        labels.showVal = True
        for flag in ("showLegendKey", "showCatName", "showSerName", "showPercent"):
            setattr(labels, flag, False)
        labels.position = "outEnd"
        labels.txPr = _text(900, brand.ink, bold=False)
        if value_format:
            labels.numFmt = value_format
        series.dLbls = labels


def _style_round(chart: Any, brand: Brand) -> None:
    chart.legend = None
    series = chart.series[0]
    for index in range(_point_count(series)):
        point = DataPoint(idx=index)
        colour = brand.palette[index % len(brand.palette)]
        point.graphicalProperties = GraphicalProperties(
            solidFill=colour, ln=LineProperties(solidFill="FFFFFF", w=RULE_EMU)
        )
        series.dPt.append(point)
    labels = DataLabelList()
    labels.showCatName, labels.showPercent = True, True
    for flag in ("showVal", "showLegendKey", "showSerName"):
        setattr(labels, flag, False)
    labels.showLeaderLines = True
    labels.separator = "\n"
    labels.txPr = _text(900, brand.ink, bold=False)
    series.dLbls = labels
    if hasattr(chart, "holeSize"):
        chart.holeSize = 55


def _point_count(series: Any) -> int:
    ref = series.val.numRef.f if series.val is not None and series.val.numRef else ""
    if not ref or ":" not in ref:
        return 1 if ref else 0
    _sheet, address = split_sheet(ref)
    rng = parse_range(address.replace("$", ""))
    return rng.n_rows * rng.n_cols


def _text(size: int, hex_colour: str, bold: bool) -> RichText:
    props = CharacterProperties(sz=size, b=bold, solidFill=hex_colour)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)])


def _title(text: str, brand: Brand) -> Title:
    props = CharacterProperties(sz=1200, b=True, solidFill=brand.ink)
    paragraph = Paragraph(
        pPr=ParagraphProperties(defRPr=props), r=[RegularTextRun(rPr=props, t=text)]
    )
    return Title(tx=Text(rich=RichText(p=[paragraph])), overlay=False)
