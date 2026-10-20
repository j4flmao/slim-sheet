"""`validate_formula` logic: will this formula work in the target Excel and workbook? Pure.

Builds on `formula.analyze` and adds: version compatibility (-> #NAME? in older Excel),
sheet/name existence when the workbook is known, unquoted sheet names, grid bounds
(.xls is 65,536 x 256), and concrete fix suggestions.
"""

from __future__ import annotations

import difflib
from collections.abc import Sequence
from dataclasses import dataclass

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.formula import FormulaReport, analyze
from slim_sheet.core.functions import NETWORK, VERSION_ORDER, newer_than
from slim_sheet.core.refs import (
    MAX_COL,
    MAX_ROW,
    XLS_MAX_COL,
    XLS_MAX_ROW,
    Ref,
    classify,
    max_extent,
    unquoted_sheet_with_space,
)

TARGETS = ("xls", *VERSION_ORDER[1:])
_XLS_REPLACEMENTS = {
    "IFERROR": "IF(ISERROR(x), fallback, x)",
    "SUMIFS": "SUMPRODUCT((range1=crit1)*(range2=crit2)*sum_range)",
    "COUNTIFS": "SUMPRODUCT((range1=crit1)*(range2=crit2))",
    "AVERAGEIFS": "SUMPRODUCT(...)/SUMPRODUCT(...)",
}
_MODERN_REPLACEMENTS = {
    "XLOOKUP": "INDEX(return_range, MATCH(value, lookup_range, 0))",
    "IFS": "nested IF(...)",
    "CONCAT": "CONCATENATE(...) or &",
    "TEXTJOIN": "& with explicit delimiters",
    "MAXIFS": "MAX(IF(criteria, range)) entered as an array formula",
    "MINIFS": "MIN(IF(criteria, range)) entered as an array formula",
}


@dataclass(frozen=True, slots=True)
class ValidationReport:
    formula: str
    target: str
    ok: bool
    errors: tuple[str, ...]
    warnings: tuple[str, ...]
    fixes: tuple[str, ...]


def validate(
    formula: str,
    target: str = "365",
    sheets: Sequence[str] | None = None,
    names: Sequence[str] | None = None,
) -> ValidationReport:
    target = target.strip().lower()
    if target not in TARGETS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown target Excel version '{target}'",
            hint=f"Use one of: {', '.join(TARGETS)}",
        )
    report = analyze(formula)
    errors = list(report.errors)
    # the catalog is complete, so an unknown function is a real #NAME? error here
    unknown = [w for w in report.warnings if "is not in the catalog" in w]
    errors += [
        w.replace("is not in the catalog", "is not an Excel function (#NAME?)") for w in unknown
    ]
    warnings = [w for w in report.warnings if w not in unknown]
    fixes: list[str] = []
    _check_versions(report, target, errors, fixes)
    _check_refs(report, target, sheets, names, errors, fixes)
    quoted = unquoted_sheet_with_space(report.formula)
    if quoted:
        errors.append(f"sheet name '{quoted}' has spaces but is not quoted")
        fixes.append(f"write '{quoted}'!A1 (single quotes around the sheet name)")
    fixes += _fixes_from_messages(errors, warnings)
    # a version problem that is already an error for this target needs no extra warning
    warnings = [w for w in warnings if not any(w.split(" ", 1)[0] + " needs" in e for e in errors)]
    return ValidationReport(
        report.formula,
        target,
        not errors,
        tuple(errors),
        tuple(warnings),
        tuple(dict.fromkeys(fixes)),
    )


def _check_versions(
    report: FormulaReport, target: str, errors: list[str], fixes: list[str]
) -> None:
    for call in report.calls:
        info = call.info
        if info is None:
            continue
        if newer_than(info, target):
            errors.append(f"{call.name} needs Excel {info.version}+; {target} shows #NAME?")
            replacement = _XLS_REPLACEMENTS.get(call.name) or _MODERN_REPLACEMENTS.get(call.name)
            if replacement:
                fixes.append(f"replace {call.name} with {replacement}")
        if call.name in NETWORK:
            errors.append(f"{call.name} reaches the network; slim-sheet refuses to write it")


def _check_refs(
    report: FormulaReport,
    target: str,
    sheets: Sequence[str] | None,
    names: Sequence[str] | None,
    errors: list[str],
    fixes: list[str],
) -> None:
    max_row, max_col = (XLS_MAX_ROW, XLS_MAX_COL) if target == "xls" else (MAX_ROW, MAX_COL)
    folded_sheets = {s.casefold(): s for s in sheets or ()}
    folded_names = {n.casefold() for n in names or ()}
    for raw in report.references:
        ref = classify(raw)
        if ref.kind == "external":
            errors.append(f"{raw} links to another workbook")
            continue
        _check_sheet(ref, sheets, folded_sheets, errors, fixes)
        row, col = max_extent(ref)
        if row > max_row or col > max_col:
            errors.append(f"{raw} is outside the {target} grid")
        if ref.kind == "name" and names is not None and ref.address.casefold() not in folded_names:
            errors.append(f"'{ref.address}' is not a defined name in this workbook (#NAME?)")
            close = difflib.get_close_matches(ref.address, list(names), n=1)
            if close:
                fixes.append(f"did you mean the name '{close[0]}'?")


def _check_sheet(
    ref: Ref,
    sheets: Sequence[str] | None,
    folded: dict[str, str],
    errors: list[str],
    fixes: list[str],
) -> None:
    if ref.sheet is None or sheets is None or ref.sheet.casefold() in folded:
        return
    errors.append(f"sheet '{ref.sheet}' does not exist (#REF!)")
    close = difflib.get_close_matches(ref.sheet, list(sheets), n=1, cutoff=0.5)
    if close:
        fixes.append(f"did you mean sheet '{close[0]}'?")


def _fixes_from_messages(errors: list[str], warnings: list[str]) -> list[str]:
    fixes = []
    if any("';'" in e for e in errors):
        fixes.append("use ',' between arguments")
    if any("missing ')'" in e for e in errors):
        fixes.append("add the missing ')' at the end of the innermost call")
    fixes += [
        e.split("; ", 1)[1].replace("did you mean", "rename to")
        for e in errors
        if "(#NAME?); did you mean" in e
    ]
    return fixes
