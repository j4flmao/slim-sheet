"""Clean, brand-coloured chart styling for openpyxl charts (PLAN §16.4).

openpyxl's defaults fight a clean business chart: gap width 150, legend on the right, major
gridlines, rounded corners, and axes Excel 365 hides unless `delete` is written explicitly.
Every chart slim-sheet builds goes through `style_chart`, which overrides all of them and
applies the report rules in docs/design.md: a left-aligned title that states the message,
numbers in the data's own format, category labels kept clear of negative bars, direct labels
instead of a redundant value axis, and no labels on slices too thin to read.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any

from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Font,
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
BAR_KINDS = ("column", "bar")
MAX_LABELLED_POINTS = 12
MAX_END_LABEL_POINTS = 60  # a line labels its last value while the points stay readable
MIN_LABEL_SHARE = 0.03  # slices under 3% get no label: it would overlap its neighbours
DARK_LUMINANCE = 0.45  # slice colours darker than this take white label text
INSIDE_LABEL_SHARE = 0.85  # bars this close to the longest one label inside their end
_NO_TITLE = {"", "none", "null", "nil", "n/a", "na", "-"}


@dataclass(frozen=True, slots=True)
class ChartText:
    """What the caller knows about the numbers and words around a chart."""

    title: str | None = None
    highlight: int | None = None  # category index drawn in the brand colour, others grey
    value_format: str | None = None  # Excel number format of the values
    x_format: str | None = None  # scatter only: format of the x values
    x_title: str | None = None  # the horizontal axis, whatever the chart kind
    y_title: str | None = None  # the vertical axis
    values: tuple[float, ...] = ()  # first-series values: signs and slice shares


def clean_title(text: str | None) -> str | None:
    """Models pass 'None', 'null' or '' for 'no title'; never print those on a chart."""
    if text is None or text.strip().casefold() in _NO_TITLE:
        return None
    return text.strip()


def style_chart(chart: Any, kind: str, brand: Brand, text: ChartText | None = None) -> None:
    """Strip chart junk and colour series from the brand; `highlight` greys all other bars."""
    words = text or ChartText()
    chart.roundedCorners = False
    chart.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    chart.style = None
    title = clean_title(words.title)
    if title:
        chart.title = _title(title, brand)
    if kind in ROUND_KINDS:
        _style_round(chart, kind, brand, words.values)
        return
    n_series = len(chart.series)
    _style_axes(chart, kind, brand, words)
    chart.legend = Legend(legendPos="t") if n_series > 1 else None
    for index, series in enumerate(chart.series):
        _colour_series(series, kind, brand.palette[index % len(brand.palette)])
    if kind == "bar":  # horizontal bars: first category on top, value axis at the bottom
        chart.x_axis.scaling.orientation = "maxMin"
        chart.y_axis.crosses = "max"
    if kind in BAR_KINDS:
        chart.gapWidth = 60 if n_series == 1 else 80
        chart.overlap = -10 if n_series > 1 else None
        if n_series == 1:
            _single_series_bars(chart, kind, brand, words)
    if kind == "line" and n_series == 1:
        _label_last_point(chart.series[0], brand, words.value_format)


def _style_axes(chart: Any, kind: str, brand: Brand, words: ChartText) -> None:
    axis_line = GraphicalProperties(ln=LineProperties(solidFill=brand.rule, w=RULE_EMU))
    for axis in (chart.x_axis, chart.y_axis):
        axis.delete = False  # Excel 365 hides axes when <c:delete> is missing
        axis.txPr = _text(900, brand.muted, brand.body_font)
        axis.spPr = axis_line
        axis.majorGridlines = None
        axis.title = None
    # labels sit at the low edge: a negative bar or point never runs through them
    chart.x_axis.tickLblPos = "low"
    if kind == "scatter":
        chart.y_axis.tickLblPos = "low"
        if words.x_format:
            chart.x_axis.number_format = words.x_format
    if kind not in BAR_KINDS:  # bars carry data labels instead of gridlines
        chart.y_axis.majorGridlines = _gridlines(brand)
    if words.value_format:
        chart.y_axis.number_format = axis_format(words.value_format, words.values)
    # x_title names the horizontal axis; on horizontal bars that is the value (y) axis
    across, up = (chart.y_axis, chart.x_axis) if kind == "bar" else (chart.x_axis, chart.y_axis)
    for axis, label in ((across, words.x_title), (up, words.y_title)):
        name = clean_title(label)
        if name:
            axis.title = _axis_title(name, brand)


def axis_format(number_format: str, values: Sequence[float]) -> str:
    """Axis ticks are round steps: '4,500' not '4,500.0', '25%' not '25.0%'. Data labels keep
    the decimals; only small ranges (under 10, or under 10%) keep them on the axis."""
    if not values:
        return number_format
    biggest = max(abs(v) for v in values)
    small = biggest < (0.1 if "%" in number_format else 10)
    return number_format if small else re.sub(r"(?<=0)\.0+", "", number_format)


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
            series.marker.symbol = "circle"
            series.marker.size = 6
            series.marker.graphicalProperties = GraphicalProperties(
                solidFill=hex_colour, ln=LineProperties(solidFill="FFFFFF", w=RULE_EMU)
            )
            props.line = LineProperties(noFill=True)
        return
    props.solidFill = hex_colour
    props.line = LineProperties(noFill=True)
    if hasattr(series, "invertIfNegative"):
        series.invertIfNegative = False  # Excel draws inverted negatives hollow (white)


def _single_series_bars(chart: Any, kind: str, brand: Brand, words: ChartText) -> None:
    series = chart.series[0]
    points = _point_count(series)
    highlight = words.highlight
    if highlight is not None and 0 <= highlight < points:
        series.graphicalProperties.solidFill = CONTEXT_GREY
        point = DataPoint(idx=highlight)
        point.graphicalProperties = GraphicalProperties(solidFill=brand.primary)
        series.dPt.append(point)
    if not 0 < points <= MAX_LABELLED_POINTS:
        return
    labels = _labels(brand, words.value_format, show_val=True)
    labels.position = "outEnd"
    labels.dLbl = _inside_end_labels(brand, words)
    series.dLbls = labels
    value_axis_titled = clean_title(words.x_title if kind == "bar" else words.y_title)
    if not value_axis_titled:  # every bar carries its number: the value axis is redundant
        chart.y_axis.delete = True


def _inside_end_labels(brand: Brand, words: ChartText) -> list[DataLabel]:
    """Bars within reach of the longest one carry their number inside the bar end: Excel
    leaves no room past the plot area, so an outside label there is cut off."""
    longest = max((v for v in words.values if v > 0), default=0.0)
    out: list[DataLabel] = []
    for index, value in enumerate(words.values):
        if longest and value >= INSIDE_LABEL_SHARE * longest:
            grey = words.highlight is not None and index != words.highlight
            fill = CONTEXT_GREY if grey else brand.primary
            ink = "FFFFFF" if _luminance(fill) < DARK_LUMINANCE else brand.ink
            label = DataLabel(idx=index, showVal=True, showLegendKey=False, showCatName=False,
                              showSerName=False, showPercent=False)  # fmt: skip
            label.position = "inEnd"
            label.txPr = _text(900, ink, brand.body_font, bold=True)
            if words.value_format:
                label.numFmt = words.value_format
            out.append(label)
    return out


def _label_last_point(series: Any, brand: Brand, value_format: str | None) -> None:
    """A trend line ends with its latest value, the number a reader looks for first."""
    points = _point_count(series)
    if not 1 < points <= MAX_END_LABEL_POINTS:
        return
    label = DataLabel(idx=points - 1, showVal=True, showLegendKey=False, showCatName=False,
                      showSerName=False, showPercent=False)  # fmt: skip
    label.position = "r"
    label.txPr = _text(900, brand.ink, brand.body_font, bold=True)
    if value_format:
        label.numFmt = value_format
    labels = DataLabelList(dLbl=[label])
    for flag in ("showVal", "showLegendKey", "showCatName", "showSerName", "showPercent"):
        setattr(labels, flag, False)
    series.dLbls = labels


def _style_round(chart: Any, kind: str, brand: Brand, values: Sequence[float]) -> None:
    """Pie: category and share beside each slice. Doughnut: a legend and the share on the
    ring, in white or ink to contrast with the slice. Thin slices stay unlabelled."""
    series = chart.series[0]
    count = _point_count(series)
    colours = [brand.palette[i % len(brand.palette)] for i in range(count)]
    for index, colour in enumerate(colours):
        point = DataPoint(idx=index)
        point.graphicalProperties = GraphicalProperties(
            solidFill=colour, ln=LineProperties(solidFill="FFFFFF", w=RULE_EMU)
        )
        series.dPt.append(point)
    ring = kind == "doughnut"
    labels = _labels(brand, "0%", show_val=False)
    labels.showPercent = True
    labels.showCatName = not ring
    labels.showLeaderLines = not ring
    labels.separator = "\n"
    if not ring:
        labels.position = "bestFit"
    labels.dLbl = _point_labels(colours, values, ring, brand)
    series.dLbls = labels
    chart.legend = Legend(legendPos="r") if ring else None
    if hasattr(chart, "holeSize"):
        chart.holeSize = 55


def _point_labels(
    colours: Sequence[str], values: Sequence[float], ring: bool, brand: Brand
) -> list[DataLabel]:
    total = sum(v for v in values if v > 0)
    out: list[DataLabel] = []
    for index, colour in enumerate(colours):
        share = values[index] / total if total and index < len(values) else None
        hidden = share is not None and share < MIN_LABEL_SHARE
        if not hidden and not ring:
            continue
        label = DataLabel(idx=index, showVal=False, showLegendKey=False, showSerName=False,
                          showCatName=False, showPercent=not hidden)  # fmt: skip
        if not hidden:
            ink = "FFFFFF" if _luminance(colour) < DARK_LUMINANCE else brand.ink
            label.txPr = _text(900, ink, brand.body_font, bold=True)
            label.numFmt = "0%"
        out.append(label)
    return out


def _luminance(hex_colour: str) -> float:
    red, green, blue = (int(hex_colour[i : i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def _labels(brand: Brand, number_format: str | None, show_val: bool) -> DataLabelList:
    labels = DataLabelList()
    labels.showVal = show_val
    for flag in ("showLegendKey", "showCatName", "showSerName", "showPercent"):
        setattr(labels, flag, False)
    labels.txPr = _text(900, brand.ink, brand.body_font)
    if number_format:
        labels.numFmt = number_format
    return labels


def _point_count(series: Any) -> int:
    ref = series.val.numRef.f if series.val is not None and series.val.numRef else ""
    if not ref and getattr(series, "yVal", None) is not None and series.yVal.numRef:
        ref = series.yVal.numRef.f  # scatter series keep their values in yVal
    if not ref or ":" not in ref:
        return 1 if ref else 0
    _sheet, address = split_sheet(ref)
    rng = parse_range(address.replace("$", ""))
    return rng.n_rows * rng.n_cols


def _props(size: int, hex_colour: str, typeface: str, bold: bool) -> CharacterProperties:
    return CharacterProperties(sz=size, b=bold, solidFill=hex_colour, latin=Font(typeface=typeface))


def _text(size: int, hex_colour: str, typeface: str, bold: bool = False) -> RichText:
    props = _props(size, hex_colour, typeface, bold)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)])


def _title(text: str, brand: Brand) -> Title:
    """The message title, top-left like the first line of a slide (not centred)."""
    props = _props(1200, brand.ink, brand.heading_font, bold=True)
    paragraph = Paragraph(
        pPr=ParagraphProperties(defRPr=props, algn="l"), r=[RegularTextRun(rPr=props, t=text)]
    )
    corner = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.01, y=0.01))
    return Title(tx=Text(rich=RichText(p=[paragraph])), overlay=False, layout=corner)


def _axis_title(text: str, brand: Brand) -> Title:
    props = _props(900, brand.muted, brand.body_font, bold=False)
    paragraph = Paragraph(
        pPr=ParagraphProperties(defRPr=props), r=[RegularTextRun(rPr=props, t=text)]
    )
    return Title(tx=Text(rich=RichText(p=[paragraph])), overlay=False)
