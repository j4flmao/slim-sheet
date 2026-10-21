"""Formula guard shared by every tool that writes formulas or rules (PLAN §8, §11.9). Pure.

Formulas that reach outside the workbook can exfiltrate data or run code when a person later
opens the file: web functions, code-calling functions, DDE, links to other workbooks and
hyperlinks/images pointing at a URL or file. String literals are ignored when matching, so
text that merely mentions a function name is fine.
"""

from __future__ import annotations

import re

from slim_sheet.core.errors import ErrorCode, SlimError

_STRING_LITERAL = re.compile(r'"[^"]*"')
_DANGEROUS = re.compile(
    r"\b(WEBSERVICE|FILTERXML|CALL|REGISTER(\.ID)?|EXEC|RTD|IMPORTXML|IMPORTDATA|IMPORTHTML)\s*\("
    r"|\[[^\]]*\.(xl\w*|ods|csv)\]|\[\d+\][^!]*!"  # links to other workbooks
    r"|\|",  # DDE, e.g. =cmd|'/c calc'!A0
    re.IGNORECASE,
)
_OUTBOUND = re.compile(r"\b(HYPERLINK|IMAGE)\s*\(\s*\"(https?:|ftp:|file:|\\\\)", re.IGNORECASE)


def outside_reason(formula: str) -> str | None:
    """Why a formula reaches outside the workbook, or None when it stays inside."""
    if _OUTBOUND.search(formula):
        return "HYPERLINK/IMAGE to an external address"
    if _DANGEROUS.search(_STRING_LITERAL.sub('""', formula)):
        return "external link, DDE or web/code function"
    return None


def check_formula(formula: str, where: str = "") -> None:
    """Raise formula_denied for a formula that reaches outside the workbook."""
    reason = outside_reason(formula)
    if reason is None:
        return
    place = f" at {where}" if where else ""
    raise SlimError(
        ErrorCode.FORMULA_DENIED,
        f"Formula{place} reaches outside the workbook ({reason}): {formula[:80]}",
        hint="External links, DDE, web functions and outbound HYPERLINK/IMAGE are blocked; "
        "for a clickable web link use annotate(op='link', url='https://...').",
    )
