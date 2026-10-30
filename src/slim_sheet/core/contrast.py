"""WCAG 2.2 colour contrast for charts and text (docs/prompts/design.md §4). Pure.

- Text needs 4.5:1 against its background (3:1 for large text: 18 pt, or 14 pt bold).
- Chart marks need 3:1 against the background (WCAG 1.4.11 non-text contrast).
A brand keeps its colours; slim-sheet only orders them so the first series are readable and
falls back to the brand's ink when a text colour would be too faint.
"""

from __future__ import annotations

from collections.abc import Sequence

TEXT_RATIO, LARGE_TEXT_RATIO, MARK_RATIO = 4.5, 3.0, 3.0
LARGE_TEXT_PT, LARGE_BOLD_PT = 18.0, 14.0


def relative_luminance(hex_colour: str) -> float:
    """WCAG relative luminance of 'RRGGBB' (0 black .. 1 white)."""
    channels = [int(hex_colour[i : i + 2], 16) / 255 for i in (0, 2, 4)]
    linear = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in channels]
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]


def contrast_ratio(first: str, second: str = "FFFFFF") -> float:
    """1.0 (identical) .. 21.0 (black on white)."""
    light, dark = sorted((relative_luminance(first), relative_luminance(second)), reverse=True)
    return (light + 0.05) / (dark + 0.05)


def text_minimum(size_pt: float | None, bold: bool) -> float:
    size = size_pt or 11.0
    large = size >= LARGE_TEXT_PT or (bold and size >= LARGE_BOLD_PT)
    return LARGE_TEXT_RATIO if large else TEXT_RATIO


def readable_text(
    colour: str, background: str, fallback: str, size_pt: float | None = None, bold: bool = False
) -> str:
    """`colour` when it is readable on `background`, else `fallback` (the brand's ink)."""
    if contrast_ratio(colour, background) >= text_minimum(size_pt, bold):
        return colour
    return fallback


def accessible_order(palette: Sequence[str], background: str = "FFFFFF") -> tuple[str, ...]:
    """Series colours that pass 3:1 first (in brand order), the faint ones last."""
    passing = [c for c in palette if contrast_ratio(c, background) >= MARK_RATIO]
    faint = [c for c in palette if contrast_ratio(c, background) < MARK_RATIO]
    return (*passing, *faint)


def label_ink(fill: str, dark: str, light: str = "FFFFFF") -> str:
    """The label colour (dark ink or white) with the higher contrast on a bar or slice."""
    return light if contrast_ratio(light, fill) >= contrast_ratio(dark, fill) else dark


def faint_series(palette: Sequence[str], count: int, background: str = "FFFFFF") -> list[str]:
    """Colours among the first `count` series that fail 3:1 against the background."""
    ordered = accessible_order(palette, background)
    used = [ordered[i % len(ordered)] for i in range(count)] if ordered else []
    return [c for c in dict.fromkeys(used) if contrast_ratio(c, background) < MARK_RATIO]
