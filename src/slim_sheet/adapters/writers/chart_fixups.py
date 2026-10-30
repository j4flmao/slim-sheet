"""Repair what openpyxl writes wrong in chart parts, after every save (PLAN §16.4).

- A data-label number format is written as <numFmt formatCode=".."/> without sourceLinked, and
  Excel then shows the source cell's format instead (0.2497, not +25.0%). Every numFmt that
  lacks the attribute was set on purpose, so it gets sourceLinked="0".
- Text properties (<txPr>) get an empty run <a:r><a:t/></a:r>; read back, the empty text
  becomes None and the next save writes the word "None". txPr holds formatting only, so its
  runs are removed.
- Every chart without alt text gets one from its kind and title (chart_alt.py).
- Excel 365 hides an axis that has no <delete>. openpyxl writes none for axes it rebuilt on
  load (an area chart loses its axis settings that way), so a missing one becomes val="0".
"""

from __future__ import annotations

import contextlib
import os
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path

from slim_sheet.adapters.writers.chart_alt import chart_targets, with_alt_text

_CHART_PART = re.compile(r"^xl/charts/chart\d+\.xml$")
_DRAWING_PART = re.compile(r"^xl/drawings/drawing\d+\.xml$")
_NUM_FMT = re.compile(r"<((?:\w+:)?numFmt)\b(?![^>]*\bsourceLinked=)([^>]*?)\s*/>")
_TX_PR = re.compile(r"<((?:\w+:)?txPr)>(.*?)</\1>", re.DOTALL)
_RUN = re.compile(r"<a:r>.*?</a:r>", re.DOTALL)
_AXIS = re.compile(r"<(\w+:)?(catAx|valAx|dateAx|serAx)>(.*?)</\1?\2>", re.DOTALL)
_SCALING_END = re.compile(r"</(?:\w+:)?scaling>|<(?:\w+:)?scaling\s*/>")
_DELETE = re.compile(r"<(?:\w+:)?delete\b")


def fix_chart_xml(xml: str) -> str:
    xml = _NUM_FMT.sub(r'<\1\2 sourceLinked="0"/>', xml)
    xml = _TX_PR.sub(lambda m: f"<{m.group(1)}>{_RUN.sub('', m.group(2))}</{m.group(1)}>", xml)
    return _AXIS.sub(_visible_axis, xml)


def _visible_axis(match: re.Match[str]) -> str:
    prefix, tag, body = match.group(1) or "", match.group(2), match.group(3)
    end = _SCALING_END.search(body)
    if _DELETE.search(body) or end is None:
        return match.group(0)
    body = f'{body[: end.end()]}<{prefix}delete val="0"/>{body[end.end() :]}'
    return f"<{prefix}{tag}>{body}</{prefix}{tag}>"


def fix_charts(path: Path) -> int:
    """Rewrite the package when a chart part needs a repair or a chart lacks alt text;
    returns the parts changed."""
    fixed: dict[str, str] = {}
    with zipfile.ZipFile(path) as source:
        names = source.namelist()
        charts = {n: source.read(n).decode("utf-8") for n in names if _CHART_PART.match(n)}
        for name, old in charts.items():
            new = fix_chart_xml(old)
            if new != old:
                fixed[name] = charts[name] = new
        for name in filter(_DRAWING_PART.match, names):
            rels = posixpath.join(posixpath.dirname(name), "_rels", posixpath.basename(name))
            if f"{rels}.rels" not in names:
                continue
            targets = chart_targets(name, source.read(f"{rels}.rels").decode("utf-8"))
            linked = {rid: charts[part] for rid, part in targets.items() if part in charts}
            old = source.read(name).decode("utf-8")
            new = with_alt_text(old, linked) if linked else old
            if new != old:
                fixed[name] = new
    if not fixed:
        return 0
    fd, tmp_name = tempfile.mkstemp(prefix=".slim-", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with (
            zipfile.ZipFile(path) as source,
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out,
        ):
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename in fixed:
                    data = fixed[item.filename].encode("utf-8")
                out.writestr(item, data)
        tmp.replace(path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            tmp.unlink()
    return len(fixed)
