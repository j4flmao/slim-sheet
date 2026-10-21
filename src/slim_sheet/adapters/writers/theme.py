"""Write a brand into a workbook: Office theme XML, named cell styles, and role styling.

The theme is what makes Excel's own table styles, chart styles and "Accent" cell styles follow
the brand when a person keeps editing the file, so a company configures its brand once.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.writer.theme import theme_xml as OPENPYXL_THEME

from slim_sheet.core.design import Brand, RoleStyle, role_style

THEME_PREFIX = "slim: "
_SLOTS = ("dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6")
_NAMED_ROLES = ("title", "subtitle", "section", "header", "total", "input", "note", "kpi_value")


def theme_xml(brand: Brand) -> str:
    """openpyxl's Office theme with the brand's colours and fonts."""
    palette = [brand.palette[i % len(brand.palette)] for i in range(6)]
    colours = dict(zip(_SLOTS, (brand.ink, brand.band, *palette), strict=True))
    name = _xml_escape(f"{THEME_PREFIX}{brand.name}")
    xml = str(OPENPYXL_THEME).replace('name="Office Theme"', f'name="{name}"', 1)
    for slot, hex_value in colours.items():
        xml = _set(xml, rf'(<a:{slot}>\s*<a:srgbClr val=")[0-9A-Fa-f]{{6}}', hex_value)
    for tag, font in (("majorFont", brand.heading_font), ("minorFont", brand.body_font)):
        xml = _set(xml, rf'(<a:{tag}>\s*<a:latin typeface=")[^"]*', _xml_escape(font))
    return xml


def _set(xml: str, pattern: str, value: str) -> str:
    """Replace what follows group 1 of the first match with a literal value."""
    match = re.search(pattern, xml)
    if match is None:
        return xml
    return xml[: match.end(1)] + value + xml[match.end() :]


def _xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")


def has_custom_theme(wb: Any) -> bool:
    """True when the workbook carries a theme someone chose (not openpyxl's or ours)."""
    raw = wb.loaded_theme
    if not raw:
        return False
    text = raw.decode("utf-8", "replace") if isinstance(raw, bytes) else str(raw)
    head = text[:600]
    return f'name="{THEME_PREFIX}' not in head and "4F81BD" not in text


def apply_brand(wb: Any, brand: Brand, force: bool) -> bool:
    """Install the brand theme and named styles; keeps a company theme unless `force`."""
    if has_custom_theme(wb) and not force:
        return False
    wb.loaded_theme = theme_xml(brand).encode("utf-8")
    existing = set(wb.named_styles)
    for role in _NAMED_ROLES:
        name = f"{brand.name} {role.replace('_', ' ')}"
        if name not in existing:
            named = NamedStyle(name=name)
            _style_parts(named, role_style(role, brand), brand)
            wb.add_named_style(named)
    return True


def style_cell(cell: Any, style: RoleStyle, brand: Brand, number_format: str | None = None) -> None:
    """Apply a role on top of the cell's current look (unset parts are kept)."""
    _style_parts(cell, style, brand)
    if number_format:
        cell.number_format = number_format


def _style_parts(target: Any, style: RoleStyle, brand: Brand) -> None:
    old = target.font
    target.font = Font(
        name=brand.heading_font if style.font == "heading" else brand.body_font,
        size=style.size if style.size is not None else old.size,
        bold=style.bold if style.bold is not None else old.bold,
        italic=old.italic,
        color=f"FF{style.color}" if style.color else old.color,
    )
    if style.fill:
        target.fill = PatternFill("solid", fgColor=f"FF{style.fill}")
    if style.top or style.bottom:
        border = target.border
        target.border = Border(
            left=border.left,
            right=border.right,
            top=Side(style=style.top, color=f"FF{brand.ink}") if style.top else border.top,
            bottom=(
                Side(style=style.bottom, color=f"FF{brand.ink}") if style.bottom else border.bottom
            ),
        )
    if style.align or style.valign or style.wrap is not None:
        old_align = target.alignment
        target.alignment = Alignment(
            horizontal=style.align or old_align.horizontal,
            vertical=style.valign or old_align.vertical,
            wrap_text=style.wrap if style.wrap is not None else old_align.wrap_text,
        )


def body_font(cell: Any, brand: Brand) -> None:
    """Switch a cell to the brand body font, keeping size, bold, italic and colour."""
    old = cell.font
    if old.name == brand.body_font:
        return
    cell.font = Font(
        name=brand.body_font, size=old.size, bold=old.bold, italic=old.italic, color=old.color
    )
