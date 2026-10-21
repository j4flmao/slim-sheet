"""Static risk triage of VBA source (PLAN §11.10). Pure; nothing is ever executed.

Finds what security reviewers look for first: procedures Excel runs by itself (auto-exec),
calls that reach the system or network, indicators such as URLs and executable names,
obfuscation, and modules whose source was removed while compiled code remains ("stomping").
A clean report is triage, not proof that a macro is safe.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

AUTO_EXEC = (
    "Auto_Open", "AutoOpen", "Auto_Close", "AutoClose", "AutoExec", "Workbook_Open",
    "Workbook_Activate", "Workbook_BeforeClose", "Workbook_BeforeSave", "Workbook_SheetChange",
    "Workbook_SheetActivate", "Worksheet_Change", "Worksheet_Activate", "Worksheet_Calculate",
    "Worksheet_SelectionChange", "Worksheet_FollowHyperlink", "Document_Open",
)  # fmt: skip
SUSPICIOUS = {
    r"\bShell\s*\(?": "runs a program (Shell)",
    r"WScript\.Shell|Shell\.Application": "Windows shell object",
    r"\bCreateObject\s*\(": "creates a COM object (CreateObject)",
    r"\bGetObject\s*\(": "attaches to a COM object (GetObject)",
    r"\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\b.*\bLib\b": "calls the Windows API (Declare)",
    r"URLDownloadToFile|XMLHTTP|WinHttpRequest|InternetExplorer\.Application": "uses the network",
    r"ADODB\.Stream|SaveToFile": "writes binary files (ADODB.Stream)",
    r"\b(Kill|FileCopy|MkDir|RmDir)\b": "changes files or folders",
    r"\bOpen\b[^\n]*\bFor\s+(Output|Binary|Append)\b": "writes a file",
    r"\bEnviron\$?\s*\(": "reads environment variables",
    r"\bReg(Write|Read|Delete)\b|GetSetting|SaveSetting": "reads or writes the registry/settings",
    r"powershell|cmd(\.exe)?\s*/c|mshta|rundll32|regsvr32|certutil|bitsadmin": "system tool",
    r"\bVBProject\b|VBComponents|CodeModule": "reads or changes VBA code itself",
    r"\bSendKeys\b": "types keystrokes (SendKeys)",
    r"ExecuteExcel4Macro|\bMacScript\b|\bCallByName\b": "runs code indirectly",
}  # fmt: skip
_IOC = {
    r"https?://[^\s\"')]+": "URL",
    r"\b(?:\d{1,3}\.){3}\d{1,3}\b": "IP address",
    r"[\w\-]+\.(exe|dll|ps1|bat|cmd|vbs|js|scr|hta|jar)\b": "executable file name",
}
_OBFUSCATION = (
    r"\bChrW?\$?\s*\(",
    r"\bStrReverse\s*\(",
    r"[A-Za-z0-9+/]{120,}={0,2}",
    r"(?:&H[0-9A-F]{2}\s*,?\s*){30,}",
)
RANK = ("auto_exec", "stomped", "suspicious", "ioc", "obfuscation")


@dataclass(frozen=True, slots=True)
class MacroFinding:
    module: str
    line: int  # 1-based; 0 when the finding is about the whole module
    kind: str  # one of RANK
    detail: str
    text: str


def triage(modules: Iterable[tuple[str, str]], stomped: Iterable[str] = ()) -> list[MacroFinding]:
    """modules: (name, source). Most severe kinds first."""
    why = "compiled code but no source: may run code nobody can read"
    found = [MacroFinding(name, 0, "stomped", why, "") for name in stomped]
    for name, source in modules:
        found += _module(name, source)
    order = {kind: i for i, kind in enumerate(RANK)}
    return sorted(found, key=lambda f: (order[f.kind], f.module, f.line))


def _module(name: str, source: str) -> list[MacroFinding]:
    found: list[MacroFinding] = []
    chr_calls = 0
    for number, line in enumerate(source.splitlines(), start=1):
        code = _strip_comment(line)
        if not code.strip():
            continue
        proc = re.match(r"\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)", code, re.I)
        if proc and proc.group(1).lower() in {a.lower() for a in AUTO_EXEC}:
            found.append(
                MacroFinding(
                    name, number, "auto_exec", f"{proc.group(1)} runs automatically", line.strip()
                )
            )
        for pattern, why in SUSPICIOUS.items():
            if re.search(pattern, code, re.I):
                found.append(MacroFinding(name, number, "suspicious", why, line.strip()))
        for pattern, why in _IOC.items():
            for match in re.finditer(pattern, line, re.I):
                found.append(
                    MacroFinding(name, number, "ioc", f"{why}: {match.group(0)[:80]}", line.strip())
                )
        chr_calls += len(re.findall(_OBFUSCATION[0], code, re.I))
        if any(re.search(p, code, re.I) for p in _OBFUSCATION[1:]):
            found.append(
                MacroFinding(
                    name,
                    number,
                    "obfuscation",
                    "reversed, encoded or byte-array text",
                    line.strip()[:120],
                )
            )
    if chr_calls >= 20:
        found.append(
            MacroFinding(name, 0, "obfuscation", f"{chr_calls} Chr() calls build hidden text", "")
        )
    return found


def _strip_comment(line: str) -> str:
    """Drop a trailing ' comment (outside string literals) and Rem lines."""
    if re.match(r"\s*Rem\b", line, re.I):
        return ""
    inside = False
    for i, ch in enumerate(line):
        if ch == '"':
            inside = not inside
        elif ch == "'" and not inside:
            return line[:i]
    return line


def auto_exec_names(modules: Iterable[tuple[str, str]]) -> list[str]:
    """Names of procedures Excel runs by itself (for overview warnings)."""
    return [f"{f.module}.{f.detail.split()[0]}" for f in triage(modules) if f.kind == "auto_exec"]
