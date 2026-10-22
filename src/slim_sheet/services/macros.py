"""Read-only macro use-cases (PLAN §11.10): list modules, show source, triage risks, export.

Nothing here runs VBA. Source text is file data and is rendered inside an untrusted block.
Running macros is a recorded non-goal (§17.3): it needs Windows + Excel and is arbitrary code
execution; agents do the macro's job with range_ops/query/clean/pivot instead.
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path

from slim_sheet.adapters.sandbox import resolve_output, resolve_path
from slim_sheet.adapters.vba import VbaProject, has_macros, read_project, zone_id
from slim_sheet.config import Settings
from slim_sheet.core.budget import estimate_tokens
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import (
    AnomalyResult,
    Finding,
    MacroListing,
    MacroModuleInfo,
    MacroSource,
    OutputResult,
)
from slim_sheet.core.vba_triage import auto_exec_names, triage
from slim_sheet.services.output import journal_before_write

VBA_SHEET = "#vba"
MAX_LINE_CHARS = 400
EXTENSIONS = {"standard": ".bas", "class": ".cls", "document": ".cls"}


def has_project(path: str, settings: Settings) -> bool:
    try:
        return has_macros(resolve_path(path, settings))[0]
    except (OSError, ValueError):
        return False


def is_vba_sheet(sheet: str | None) -> bool:
    return bool(sheet) and sheet.strip().lower() == VBA_SHEET  # type: ignore[union-attr]


def _project(path: str, settings: Settings) -> tuple[Path, VbaProject]:
    source = resolve_path(path, settings)
    project = read_project(source)
    if project is None:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"'{source.name}' has no VBA project",
            hint="Only .xlsm/.xltm/.xlam/.xlsb/.xls files carry macros; see overview(path).",
        )
    return source, project


def _code(source: str) -> list[str]:
    return [line for line in source.splitlines() if not line.startswith("Attribute VB_")]


def list_macros(path: str, settings: Settings) -> MacroListing:
    source, project = _project(path, settings)
    modules = tuple(
        MacroModuleInfo(
            m.name,
            m.kind,
            sum(1 for line in _code(m.source) if line.strip()),
            tuple(auto_exec_names([(m.name, m.source)])),
        )
        for m in project.modules
    )
    zone = zone_id(source)
    return MacroListing(
        source.name, project.name, modules, project.signed, project.locked,
        project.macro_sheets, project.stomped, zone is not None and zone >= 3,
    )  # fmt: skip


def macro_source(
    path: str,
    settings: Settings,
    module: str,
    cursor: str | None = None,
    max_tokens: int | None = None,
) -> MacroSource:
    source, project = _project(path, settings)
    wanted = module.strip().casefold()
    found = next((m for m in project.modules if m.name.casefold() == wanted), None)
    if found is None:
        names = ", ".join(m.name for m in project.modules)
        raise SlimError(ErrorCode.BAD_RANGE, f"No VBA module '{module}'", hint=f"Modules: {names}")
    lines = _code(found.source)
    start = _cursor_line(cursor)
    budget = max_tokens or settings.max_tokens
    taken: list[str] = []
    used = 0
    for text in lines[start:]:
        text = text if len(text) <= MAX_LINE_CHARS else text[:MAX_LINE_CHARS] + " …"
        cost = estimate_tokens(text) + 2
        if taken and used + cost > budget:
            break
        taken.append(text)
        used += cost
    end = start + len(taken)
    nxt = f"line:{end + 1}" if end < len(lines) else None
    return MacroSource(
        source.name, found.name, found.kind, start + 1, tuple(taken), len(lines), nxt
    )


def _cursor_line(cursor: str | None) -> int:
    if not cursor:
        return 0
    text = cursor.strip().lower()
    if text.startswith("line:") and text[5:].isdigit() and int(text[5:]) >= 1:
        return int(text[5:]) - 1
    raise SlimError(
        ErrorCode.BAD_CURSOR,
        f"Bad cursor '{cursor}'",
        hint="Pass the cursor from the last reply, e.g. 'line:61'",
    )


def macro_audit(path: str, settings: Settings, max_results: int = 20) -> AnomalyResult:
    source, project = _project(path, settings)
    # numbered like read(sheet="#vba") shows them: Attribute lines are hidden there too
    code = [(m.name, "\n".join(_code(m.source))) for m in project.modules]
    found = triage(code, project.stomped)
    findings = tuple(
        Finding(f"{f.module}:{f.line}" if f.line else f.module, f.text, f.kind, f.detail)
        for f in found[: max(1, min(max_results, 200))]
    )
    counts = tuple(Counter(f.kind for f in found).items())
    lines = sum(len(_code(m.source)) for m in project.modules)
    return AnomalyResult(
        source.name, f"VBA project {project.name}", "macros", lines, findings, counts
    )


def export_macros(
    path: str, settings: Settings, out_dir: str, overwrite: bool = False
) -> OutputResult:
    """Write each module as .bas/.cls text next to the workbook (Excel: File > Import File)."""
    source, project = _project(path, settings)
    folder = resolve_output(out_dir, settings, base=source.parent)
    if folder.suffix:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "out_path for a macro export is a folder",
            hint="e.g. out_path='macros'",
        )
    folder.mkdir(parents=True, exist_ok=True)
    written = []
    for module in project.modules:
        target = folder / f"{module.name}{EXTENSIONS.get(module.kind, '.bas')}"
        if target.exists() and not overwrite:
            raise SlimError(
                ErrorCode.WRITE_CONFLICT,
                f"'{target.name}' already exists",
                hint="Pass overwrite=true.",
            )
        journal_before_write(target, f"export macro {module.name}")
        target.write_text(
            module.source.replace("\n", "\r\n"), encoding="cp1252", errors="replace", newline=""
        )
        written.append(target.name)
    summary = f"{len(written)} VBA module(s) exported to {folder.name}/: {', '.join(written)}"
    return OutputResult(
        source.name,
        folder.name,
        summary,
        None,
        ("import them in Excel: VBA editor > File > Import File",),
    )


def macro_warnings(path: Path) -> list[str]:
    """Overview warnings from the file's real content (not its name); [] when it has no macros."""
    try:
        present, sheets, signed = has_macros(path)
    except (OSError, ValueError, SlimError):
        return []
    found: list[str] = []
    if present:
        try:
            project = read_project(path)
        except SlimError as err:
            return [f"contains VBA macros that could not be decoded ({err.message})"]
        if project is not None:
            autos = auto_exec_names([(m.name, m.source) for m in project.modules])
            extra = [t for t, on in (("signed", signed), ("password-locked", project.locked)) if on]
            if autos:
                extra.append("runs on its own: " + ", ".join(autos[:3]))
            if project.stomped:
                extra.append("compiled code without source (stomping)")
            detail = f" ({'; '.join(extra)})" if extra else ""
            found.append(
                f"contains VBA macros: {len(project.modules)} module(s){detail}; slim-sheet never "
                "runs them - read(path, sheet='#vba') shows the code, "
                "anomalies(path, scope='macros') rates risk"
            )
    if sheets:
        found.append(f"has {sheets} Excel 4.0 (XLM) macro sheet(s)")
    zone = zone_id(path)
    if (present or sheets) and zone is not None and zone >= 3:
        found.append("downloaded from the internet (Mark of the Web): Excel blocks its macros")
    return found
