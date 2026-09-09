"""
overflow.py
Overflow detection and a tiered auto-fix strategy, per points #4-#5 in the
original design discussion: font size is NOT changed immediately, only when
the text genuinely overflows.

Honest caveat about the limits of this approach:
    Measuring "does this text overflow" with 100% accuracy requires actual
    rendering (PowerPoint / LibreOffice), because kerning, hinting, and
    PowerPoint's built-in text-fit-shrink can't be reproduced exactly by a
    Python-side estimate. This module uses an ESTIMATE based on Pillow
    ImageFont metrics (per-character width) plus wrapping to the box width,
    compared against the box height. For an accurate final check, still
    render to images and inspect visually (see validator.py / soffice) --
    this is the "visual regression" step from point #13 of the discussion.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from PIL import ImageFont
from pptx.text.text import _Paragraph
from pptx.util import Emu, Pt

# Fallback font for width estimation when the original font isn't available
# on the system. DejaVuSans ships preinstalled on almost every Linux distro
# (including this sandbox).
_FALLBACK_FONT = "DejaVuSans.ttf"
_FONT_CACHE: dict[tuple[str, int], ImageFont.FreeTypeFont] = {}


def _get_font(size_pt: float) -> ImageFont.FreeTypeFont:
    key = (_FALLBACK_FONT, int(size_pt))
    if key not in _FONT_CACHE:
        try:
            _FONT_CACHE[key] = ImageFont.truetype(_FALLBACK_FONT, int(size_pt))
        except OSError:
            _FONT_CACHE[key] = ImageFont.load_default()
    return _FONT_CACHE[key]


@dataclass
class FitResult:
    fits: bool
    estimated_lines: int
    box_capacity_lines: int
    font_size_used: Pt


def _wrap_line_count(text: str, font: ImageFont.FreeTypeFont, max_width_px: float) -> int:
    """Estimate how many lines `text` would wrap into at width max_width_px."""
    if not text:
        return 1
    words = text.split()
    lines = 1
    cur = ""
    for w in words:
        trial = f"{cur} {w}".strip()
        w_px = font.getlength(trial)
        if w_px <= max_width_px:
            cur = trial
        else:
            lines += 1
            cur = w
    return max(lines, 1)


def estimate_fit(
    paragraph: _Paragraph,
    box_width_emu: int,
    box_height_emu: int,
    default_font_size_pt: float = 18.0,
    line_spacing_factor: float = 1.2,
    internal_margin_emu: int = 91440,  # ~0.1" default PowerPoint text inset
) -> FitResult:
    """
    Estimate whether `paragraph` (after its text is replaced) fits inside
    the box. box_width_emu / box_height_emu: shape dimensions in EMU
    (shape.width / shape.height).
    """
    text = paragraph.text
    font_size_pt = default_font_size_pt
    for run in paragraph.runs:
        if run.font.size is not None:
            font_size_pt = run.font.size.pt
            break

    usable_width_emu = max(box_width_emu - 2 * internal_margin_emu, 1)
    # 1pt = 12700 EMU; px estimate at 96dpi: 1pt = 96/72 px
    usable_width_px = Emu(usable_width_emu).pt * (96 / 72)

    font = _get_font(font_size_pt)
    lines_needed = 0
    for line in text.split("\n"):
        lines_needed += _wrap_line_count(line, font, usable_width_px)

    line_height_pt = font_size_pt * line_spacing_factor
    usable_height_emu = max(box_height_emu - 2 * internal_margin_emu, 1)
    usable_height_pt = Emu(usable_height_emu).pt
    capacity_lines = max(int(usable_height_pt // line_height_pt), 1)

    return FitResult(
        fits=lines_needed <= capacity_lines,
        estimated_lines=lines_needed,
        box_capacity_lines=capacity_lines,
        font_size_used=Pt(font_size_pt),
    )


def autofit_paragraph(
    shape,
    paragraph: _Paragraph,
    min_font_size_pt: float = 10.0,
    step_pt: float = 1.0,
) -> dict:
    """
    Apply the tiered strategy from the discussion:
      Level 1: leave it alone (original size).
      Level 2: tighten character spacing slightly (-0.3pt) -- via XML rPr@spc.
      Level 3: step the font size down until it fits or hits the minimum.
      Level 4: tighten line spacing (1.2 -> 1.05).
      Level 5: NOT done automatically here (enlarging the box could collide
               with other elements) -- returned as a 'needs_manual_resize'
               flag so it can be reviewed by a human / handled in a
               separate layer.

    Returns a report dict used for the summary shown to the user (point #11
    in the discussion: "3 text boxes resized, 1 font size reduced").
    """
    report = {
        "location": None,
        "action": "none",
        "original_size_pt": None,
        "final_size_pt": None,
        "needs_manual_resize": False,
    }

    box_w = shape.width
    box_h = shape.height

    fit = estimate_fit(paragraph, box_w, box_h)
    report["original_size_pt"] = fit.font_size_used.pt
    if fit.fits:
        return report  # Level 1: nothing needed

    # Level 2: tighten character spacing a bit
    _set_char_spacing(paragraph, -0.3)
    fit = estimate_fit(paragraph, box_w, box_h)
    if fit.fits:
        report["action"] = "tightened_char_spacing"
        return report

    # Level 3: step the font size down
    current_size = fit.font_size_used.pt
    while current_size - step_pt >= min_font_size_pt:
        current_size -= step_pt
        _set_font_size(paragraph, current_size)
        fit = estimate_fit(paragraph, box_w, box_h, default_font_size_pt=current_size)
        if fit.fits:
            report["action"] = "reduced_font_size"
            report["final_size_pt"] = current_size
            return report

    # Level 4: tighten line spacing as a last automatic attempt
    paragraph.line_spacing = 1.05
    fit = estimate_fit(paragraph, box_w, box_h, default_font_size_pt=current_size,
                        line_spacing_factor=1.05)
    if fit.fits:
        report["action"] = "reduced_line_spacing"
        report["final_size_pt"] = current_size
        return report

    # Level 5: still overflowing -> flag for manual review, don't auto-resize the box
    report["action"] = "still_overflowing"
    report["needs_manual_resize"] = True
    report["final_size_pt"] = current_size
    return report


def _set_font_size(paragraph: _Paragraph, size_pt: float) -> None:
    for run in paragraph.runs:
        run.font.size = Pt(size_pt)


def _set_char_spacing(paragraph: _Paragraph, spacing_pt: float) -> None:
    """
    python-pptx has no official API for character spacing (rPr@spc), so we
    set it directly on the XML. spc is in units of 1/100 pt (can be negative).
    """
    spc_val = str(int(spacing_pt * 100))
    for run in paragraph.runs:
        rPr = run._r.get_or_add_rPr()
        rPr.set("spc", spc_val)
