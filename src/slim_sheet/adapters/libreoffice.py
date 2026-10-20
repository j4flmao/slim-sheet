"""Headless LibreOffice recalculation (PLAN §5.3 `recalc`).

LibreOffice does not recalculate xlsx formulas on load by default. We start it with a
throwaway user profile whose registry forces "always recalculate", convert the file to
the same format, and the output then carries fresh cached values for every formula.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from slim_sheet.core.errors import ErrorCode, SlimError

TIMEOUT_SECONDS = 180
_CANDIDATES = (
    "soffice",
    "libreoffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
)
_FILTERS = {".xlsx": "xlsx:Calc MS Excel 2007 XML", ".xlsm": "xlsm:Calc MS Excel 2007 VBA XML"}
_RECALC_ITEM = (
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
    '<prop oor:name="{name}" oor:op="fuse"><value>0</value></prop></item>'
)
_RECALC_ALWAYS = "\n".join(
    [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
        'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
        _RECALC_ITEM.format(name="OOXMLRecalcMode"),  # value 0 = always recalculate on load
        _RECALC_ITEM.format(name="ODFRecalcMode"),
        "</oor:items>",
        "",
    ]
)
INSTALL_HINT = (
    "Install LibreOffice (https://www.libreoffice.org) or set SLIM_SHEET_SOFFICE to soffice. "
    "For single cells, evaluate(path=..., cell=...) works without it."
)


def find_soffice(configured: str | None) -> str:
    for candidate in (configured,) if configured else _CANDIDATES:
        found = shutil.which(candidate) or (candidate if Path(candidate).is_file() else None)
        if found:
            return found
    raise SlimError(
        ErrorCode.MISSING_DEPENDENCY, "LibreOffice (soffice) was not found", hint=INSTALL_HINT
    )


def write_profile(root: Path) -> Path:
    """Create a LibreOffice user profile that recalculates every formula on load."""
    user = root / "user"
    user.mkdir(parents=True, exist_ok=True)
    (user / "registrymodifications.xcu").write_text(_RECALC_ALWAYS, encoding="utf-8")
    return root


def build_command(soffice: str, profile: Path, source: Path, out_dir: Path) -> list[str]:
    filter_name = _FILTERS.get(source.suffix.lower())
    if filter_name is None:
        raise SlimError(
            ErrorCode.UNSUPPORTED_FORMAT, f"recalc supports .xlsx/.xlsm, not '{source.name}'"
        )
    return [
        soffice,
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--convert-to",
        filter_name,
        "--outdir",
        str(out_dir),
        str(source),
    ]


def recalculate(source: Path, soffice: str) -> Path:
    """Recalculate `source` into a temporary file and return its path (caller moves it)."""
    work = Path(tempfile.mkdtemp(prefix="slim-recalc-"))
    profile = write_profile(work / "profile")
    out_dir = work / "out"
    out_dir.mkdir()
    command = build_command(soffice, profile, source, out_dir)
    try:
        done = subprocess.run(
            command, capture_output=True, text=True, timeout=TIMEOUT_SECONDS, check=False
        )
    except subprocess.TimeoutExpired as exc:
        raise SlimError(
            ErrorCode.READ_FAILED,
            f"LibreOffice timed out after {TIMEOUT_SECONDS}s on {source.name}",
        ) from exc
    result = out_dir / source.name
    if done.returncode != 0 or not result.is_file():
        detail = (done.stderr or done.stdout).strip().splitlines()[-1:] or ["no output"]
        raise SlimError(
            ErrorCode.READ_FAILED, f"LibreOffice could not recalculate {source.name}: {detail[0]}"
        )
    return result
