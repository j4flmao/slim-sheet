"""Excel function catalog for formula validation (PLAN §10).

The data lives in `slim_sheet/data/functions.psv` (one row per function, '# ---- <category>'
headers). Reading bundled package data at import is static configuration, not user I/O.
"""

from __future__ import annotations

from dataclasses import dataclass
from importlib import resources

VARIADIC = -1
MODERN_VERSIONS = {"2019", "2021", "365"}
VERSION_ORDER = ("2003", "2007", "2010", "2013", "2016", "2019", "2021", "365")
VOLATILE = {
    "NOW",
    "TODAY",
    "RAND",
    "RANDBETWEEN",
    "RANDARRAY",
    "OFFSET",
    "INDIRECT",
    "CELL",
    "INFO",
}
NETWORK = {"WEBSERVICE", "FILTERXML", "STOCKHISTORY", "IMAGE", "TRANSLATE", "RTD"}


@dataclass(frozen=True, slots=True)
class FunctionInfo:
    name: str
    min_args: int
    max_args: int
    signature: str
    summary: str
    since: str  # "" = Excel 2003 / .xls
    category: str

    @property
    def version(self) -> str:
        return self.since or "2003"


def _parse(text: str) -> dict[str, FunctionInfo]:
    catalog: dict[str, FunctionInfo] = {}
    category = "other"
    for raw in text.splitlines():
        line = raw.strip()
        if line.startswith("# ----"):
            category = line.removeprefix("# ----").strip()
            continue
        if not line or line.startswith("#"):
            continue
        name, lo, hi, signature, summary, since = line.split("|")
        catalog[name] = FunctionInfo(name, int(lo), int(hi), signature, summary, since, category)
    return catalog


def _load() -> dict[str, FunctionInfo]:
    data = resources.files("slim_sheet.data").joinpath("functions.psv")
    return _parse(data.read_text(encoding="utf-8"))


FUNCTIONS: dict[str, FunctionInfo] = _load()


def newer_than(info: FunctionInfo, target: str) -> bool:
    """True if `info` is unavailable in Excel version `target` (e.g. '2016', 'xls')."""
    target = "2003" if target.lower() in {"xls", "2003"} else target
    if target not in VERSION_ORDER:
        return False
    return VERSION_ORDER.index(info.version) > VERSION_ORDER.index(target)
