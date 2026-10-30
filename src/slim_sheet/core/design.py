"""Design system: brands, profiles, number formats and semantic cell roles (PLAN §16). Pure.

The model chooses names (a brand, a preset, a role, a chart intent) and never hex codes or
format strings: everything visual is resolved here, deterministically, from a `Brand`.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, replace

from slim_sheet.core.contrast import readable_text
from slim_sheet.core.errors import ErrorCode, SlimError

PROFILES = ("finance", "ibcs", "gov")
NEGATIVES = ("parens", "minus")
# Okabe-Ito, reordered so the first four pass WCAG 3:1 on white (orange and sky blue do not)
OKABE_ITO = ("0072B2", "D55E00", "009E73", "CC79A7", "E69F00", "56B4E9")
CONTEXT_GREY = "BFBFBF"  # non-highlighted bars and other context ink in charts
DASH = chr(0x2013)  # en dash shown for zero in finance formats


@dataclass(frozen=True, slots=True)
class Brand:
    """Everything a workbook needs to look like one company made it."""

    name: str
    primary: str  # hex RRGGBB, no '#'
    palette: tuple[str, ...]  # chart series colours, primary first
    heading_font: str
    body_font: str
    ink: str = "262626"
    muted: str = "595959"
    rule: str = "D9D9D9"
    band: str = "F2F2F2"
    good: str = "326405"
    bad: str = "B30000"
    profile: str = "finance"
    negatives: str = "parens"
    currency: str = ""  # symbol shown by the 'currency' format, e.g. '$' or '€'
    locale: str = "en"  # 'vi': Vietnamese KPI wording and dd/mm/yyyy dates (brand 'vn')


BRANDS: dict[str, Brand] = {
    "slim": Brand("slim", "156082", ("156082", *OKABE_ITO[1:]), "Calibri Light", "Calibri"),
    "office": Brand(
        "office",
        "4472C4",
        ("4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"),
        "Calibri Light",
        "Calibri",
    ),
    "modern": Brand(
        "modern",
        "156082",
        ("156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E"),
        "Aptos Display",
        "Aptos Narrow",
    ),
    "mono": Brand(
        "mono",
        "404040",
        ("404040", "A6A6A6", "7F7F7F", "D9D9D9", "262626", "BFBFBF"),
        "Arial",
        "Arial",
        profile="ibcs",
        negatives="minus",
    ),
    "gov": Brand(
        "gov",
        "1D70B8",
        ("1D70B8", "F46A25", "28A197", "801650", "12436D", "A285D1"),
        "Arial",
        "Arial",
        ink="0B0C0C",
        muted="505A5F",
        good="326405",
        bad="B30000",
        profile="gov",
        negatives="minus",
    ),
}
BRANDS["vn"] = replace(BRANDS["slim"], name="vn", currency="", locale="vi")
DEFAULT_BRAND = "slim"
LOCALES = ("en", "vi")
LANGUAGES = {"en": "en-US", "vi": "vi-VN"}
_BRAND_KEYS = {
    "primary",
    "palette",
    "heading_font",
    "body_font",
    "ink",
    "muted",
    "rule",
    "band",
    "good",
    "bad",
    "profile",
    "negatives",
    "currency",
    "locale",
}


def builtin_brand(name: str) -> Brand:
    key = name.strip().lower()
    if key not in BRANDS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown brand '{name}'",
            hint=f"Use one of: {', '.join(BRANDS)}, or a path to a brand .json file",
        )
    return BRANDS[key]


def brand_from_mapping(data: Mapping[str, object], source: str) -> Brand:
    """A brand.json object -> Brand. Only `name` is required; the rest layers over 'base'."""
    name = str(data.get("name") or "").strip()
    if not name:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Brand file {source} needs a 'name'",
            hint='e.g. {"name": "Acme", "primary": "#C8102E", "body_font": "Arial"}',
        )
    base = builtin_brand(str(data.get("base") or DEFAULT_BRAND))
    unknown = sorted(set(data) - _BRAND_KEYS - {"name", "base"})
    if unknown:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Brand file {source} has unknown key(s): {', '.join(unknown)}",
            hint=f"Allowed: name, base, {', '.join(sorted(_BRAND_KEYS))}",
        )
    changes: dict[str, object] = {"name": name}
    for key in _BRAND_KEYS & set(data):
        changes[key] = _brand_value(key, data[key], source)
    brand = replace(base, **changes)  # type: ignore[arg-type]
    if "primary" in changes and "palette" not in changes:
        brand = replace(brand, palette=(brand.primary, *OKABE_ITO[1:]))
    return brand


def _brand_value(key: str, value: object, source: str) -> object:
    if key == "palette":
        if not isinstance(value, list) or not value:
            raise SlimError(ErrorCode.BAD_ARGUMENT, f"{source}: 'palette' must be a list of hex")
        return tuple(hex_color(str(v), f"{source} palette") for v in value)
    if key in ("profile", "negatives", "locale"):
        options = {"profile": PROFILES, "negatives": NEGATIVES, "locale": LOCALES}[key]
        text = str(value).strip().lower()
        if text not in options:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"{source}: {key} '{value}' is not valid",
                hint=f"Use one of: {', '.join(options)}",
            )
        return text
    if key in ("heading_font", "body_font", "currency"):
        return str(value).strip()
    return hex_color(str(value), f"{source} {key}")


def hex_color(value: str, label: str) -> str:
    text = value.strip().lstrip("#").upper()
    if len(text) != 6 or any(c not in "0123456789ABCDEF" for c in text):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, f"{label}: '{value}' is not a #RRGGBB colour", hint="#C8102E"
        )
    return text


# --- number formats -----------------------------------------------------------------------

FORMAT_NAMES = (
    "integer",
    "number1",
    "number2",
    "currency",
    "percent",
    "growth",
    "multiple",
    "k",
    "m",
    "bn",
    "date",
    "year",
    "text",
)


def _signed(positive: str, brand: Brand) -> str:
    """'#,##0' -> finance '#,##0_);(#,##0);"<dash>"_)' or plain minus formats for ibcs/gov."""
    if brand.negatives == "minus":
        return positive
    zero = '"0"' if brand.profile == "gov" else f'"{DASH}"'
    return f"{positive}_);({positive});{zero}_)"


def formats(brand: Brand) -> dict[str, str]:
    """Named number formats for this brand's profile and sign convention."""
    symbol = f'"{brand.currency}"' if brand.currency else ""
    return {
        "integer": _signed("#,##0", brand),
        "number1": _signed("#,##0.0", brand),
        "number2": _signed("#,##0.00", brand),
        "currency": _signed(f"{symbol}#,##0", brand),
        "percent": _signed("0.0%", brand),
        "growth": "+0.0%;-0.0%;0.0%",
        "multiple": '0.0"x"',
        "k": _signed('#,##0,"k"', brand),
        "m": _signed('#,##0.0,,"m"', brand),
        "bn": _signed('#,##0.0,,,"bn"', brand),
        "date": _date_format(brand),
        "year": "0",
        "text": "@",
    }


def _date_format(brand: Brand) -> str:
    if brand.locale == "vi":
        return "dd/mm/yyyy"
    return "d mmm yyyy" if brand.profile == "gov" else "yyyy-mm-dd"


def scaled_format(key: str, magnitude: float) -> str:
    """Pick a compact format for a single headline number (KPI tile)."""
    if key not in ("integer", "number1", "number2", "currency"):
        return key
    size = abs(magnitude)
    if size >= 1e9:
        return "bn"
    if size >= 1e6:
        return "m"
    if size >= 1e5:
        return "k"
    return key


# --- semantic roles -----------------------------------------------------------------------

ROLES = (
    "title",
    "subtitle",
    "section",
    "header",
    "total",
    "input",
    "formula",
    "link",
    "external",
    "check",
    "kpi_label",
    "kpi_value",
    "kpi_note",
    "note",
    "good",
    "bad",
)


@dataclass(frozen=True, slots=True)
class RoleStyle:
    size: float | None = None
    bold: bool | None = None
    color: str | None = None
    fill: str | None = None
    top: str | None = None  # border style for the top edge: thin | double
    bottom: str | None = None
    align: str | None = None
    valign: str | None = None
    wrap: bool | None = None
    font: str | None = None  # 'heading' uses the brand's heading font


def role_style(role: str, brand: Brand) -> RoleStyle:
    key = role.strip().lower()
    gov = brand.profile == "gov"

    def accent(background: str, size: float) -> str:  # a faint brand colour falls back to ink
        return readable_text(brand.primary, background, brand.ink, size, bold=True)

    styles = {
        "title": RoleStyle(16, True, brand.ink if gov else accent("FFFFFF", 16), font="heading"),
        "subtitle": RoleStyle(10, False, brand.muted),
        "section": RoleStyle(12, True, brand.ink, bottom="thin", font="heading"),
        "header": RoleStyle(
            None, True, brand.ink, None if gov else brand.band, bottom="thin", wrap=True,
            valign="bottom",
        ),
        "total": RoleStyle(
            None, True, brand.ink, top="thin", bottom="double" if brand.profile == "finance"
            else None,
        ),
        "input": RoleStyle(color="0000FF", fill=None if gov else "FFF2CC"),
        "formula": RoleStyle(color="000000"),
        "link": RoleStyle(color="326405"),
        "external": RoleStyle(color="B30000"),
        "check": RoleStyle(bold=True, color="B30000"),
        "kpi_label": RoleStyle(10, True, brand.muted, align="center"),
        "kpi_value": RoleStyle(24, True, brand.ink if gov else accent(brand.band, 24),
                               align="center"),
        "kpi_note": RoleStyle(9, False, brand.muted, align="center"),
        "note": RoleStyle(10 if gov else 9, False, brand.muted),
        "good": RoleStyle(bold=True, color=brand.good),
        "bad": RoleStyle(bold=True, color=brand.bad),
    }  # fmt: skip
    if key not in styles:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown role '{role}'",
            hint=f"Use one of: {', '.join(ROLES)}",
        )
    return styles[key]
