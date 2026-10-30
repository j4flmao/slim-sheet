"""Alt text for every chart (WCAG 1.1.1; Excel's Accessibility Checker reports its absence as
an error). Pure functions over the XML parts, used by `chart_fixups` after each save.

The text comes from the chart itself: its kind, its title (slim-sheet titles state the
message) and how many series it draws, at most 250 characters (Power BI's limit). A chart that
already has a description, written by a person or by Excel, keeps it.
"""

from __future__ import annotations

import posixpath
import re
from xml.sax.saxutils import quoteattr

MAX_ALT_CHARS = 250
_FRAME = re.compile(r"<((?:\w+:)?)graphicFrame\b.*?</\1graphicFrame>", re.DOTALL)
_CNVPR = re.compile(r"<((?:\w+:)?)cNvPr\b([^>]*?)(\s*/?)>")
_CHART_REF = re.compile(r"<(?:\w+:)?chart\b[^>]*\br:id=\"([^\"]+)\"")
_REL = re.compile(r"<Relationship\b[^>]*>")
_ATTR = re.compile(r'(\w+)="([^"]*)"')
_TITLE = re.compile(r"<(?:\w+:)?title>(.*?)</(?:\w+:)?title>", re.DOTALL)
_PARAGRAPH = re.compile(r"<a:p>(.*?)</a:p>", re.DOTALL)
_TEXT = re.compile(r"<a:t>([^<]*)</a:t>")
_SERIES = re.compile(r"<(?:\w+:)?ser>")
_KINDS = (
    ("doughnutChart", "Doughnut"),
    ("pieChart", "Pie"),
    ("lineChart", "Line"),
    ("areaChart", "Area"),
    ("scatterChart", "Scatter"),
    ("radarChart", "Radar"),
    ("bubbleChart", "Bubble"),
)


def describe_chart(xml: str) -> str:
    """'Bar chart: North leads Revenue with 41% of the total (3 series)'."""
    head = xml.split("plotArea>", 1)[0]  # the chart title comes before the plot area
    found = _TITLE.search(head)
    lines = []
    if found:
        for paragraph in _PARAGRAPH.findall(found.group(1)):
            text = "".join(_TEXT.findall(paragraph)).strip()
            if text:
                lines.append(text)
    series = len(_SERIES.findall(xml))
    kind = _kind(xml)
    if not lines:
        return f"{kind} chart with {series} series"[:MAX_ALT_CHARS]
    suffix = f" ({series} series)" if series > 1 else ""
    return f"{kind} chart: {' - '.join(lines)}{suffix}"[:MAX_ALT_CHARS]


def _kind(xml: str) -> str:
    if re.search(r"<(?:\w+:)?barChart\b", xml):
        return "Bar" if re.search(r'<(?:\w+:)?barDir val="bar"', xml) else "Column"
    for tag, name in _KINDS:
        if re.search(rf"<(?:\w+:)?{tag}\b", xml):
            return name
    return "Chart"


def chart_targets(drawing_part: str, rels_xml: str) -> dict[str, str]:
    """rId -> package part name of each chart the drawing links to."""
    return rel_targets(drawing_part, rels_xml, "/chart")


def rel_targets(part: str, rels_xml: str, kind: str) -> dict[str, str]:
    """rId -> package part name of every relationship of type `kind` ('/chart', '/drawing')
    that `part` (the owner of `rels_xml`) declares."""
    folder = posixpath.dirname(part)
    targets: dict[str, str] = {}
    for tag in _REL.findall(rels_xml):
        attrs = dict(_ATTR.findall(tag))
        if not attrs.get("Type", "").endswith(kind) or "Id" not in attrs:
            continue
        target = attrs.get("Target", "")
        name = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join(folder, target))
        )
        targets[attrs["Id"]] = name
    return targets


def with_alt_text(drawing_xml: str, charts: dict[str, str]) -> str:
    """Add descr="..." to every chart frame that has none; `charts` maps rId -> chart XML."""

    def frame(match: re.Match[str]) -> str:
        block = match.group(0)
        ref = _CHART_REF.search(block)
        props = _CNVPR.search(block)
        if ref is None or props is None or " descr=" in props.group(2):
            return block
        chart = charts.get(ref.group(1))
        if chart is None:
            return block
        alt = quoteattr(describe_chart(chart))
        tag = f"<{props.group(1)}cNvPr{props.group(2)} descr={alt}{props.group(3)}>"
        return block[: props.start()] + tag + block[props.end() :]

    return _FRAME.sub(frame, drawing_xml)
