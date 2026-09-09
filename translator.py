"""
translator.py
Main orchestrator. This implements the "MVP" pipeline from the original
design discussion:

    PPTX -> Parser -> Extract Text -> Translation Engine -> Replace Text
          -> Layout Checker -> (FIT | OVERFLOW -> Font adjustment)
          -> Save PPTX

The translation engine is pluggable and swappable -- see translate_engine.py
for the paid (Anthropic API) and free (Argos / Google-unofficial) options.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pptx.enum.shapes import MSO_SHAPE_TYPE

from .core import TextUnit, extract_text_units, open_presentation, save_presentation
from .overflow import autofit_paragraph
from .translate_engine import Engine, build_engine, translate_units


@dataclass
class TranslationReport:
    slides_translated: int = 0
    text_objects: int = 0
    tables_cells: int = 0
    font_reduced: int = 0
    line_spacing_reduced: int = 0
    char_spacing_tightened: int = 0
    still_overflowing: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "slides_translated": self.slides_translated,
            "text_objects": self.text_objects,
            "tables_cells": self.tables_cells,
            "layout_adjustments": {
                "font_size_reduced": self.font_reduced,
                "line_spacing_reduced": self.line_spacing_reduced,
                "char_spacing_tightened": self.char_spacing_tightened,
            },
            "still_overflowing": self.still_overflowing,
        }


def translate_pptx(
    input_path: str,
    output_path: str,
    engine: str | Engine = "argos",
    engine_kwargs: dict | None = None,
    include_notes: bool = True,
    autofit: bool = True,
) -> TranslationReport:
    """
    engine: either a string naming a built-in backend ("anthropic", "argos",
            "google_free" -- see translate_engine.py), or an already-built
            Engine instance if you want full control (e.g. a custom API key,
            a different model, or your own backend).
    engine_kwargs: extra kwargs forwarded to the engine constructor when
            `engine` is a string, e.g. {"model": "claude-sonnet-4-6"}.
    """
    if isinstance(engine, str):
        engine = build_engine(engine, **(engine_kwargs or {}))

    prs = open_presentation(input_path)

    units = extract_text_units(prs, include_notes=include_notes)
    translations = translate_units(units, engine)

    report = TranslationReport()
    report.slides_translated = len({u.slide_index for u in units})
    report.text_objects = sum(1 for u in units if u.kind == "shape")
    report.tables_cells = sum(1 for u in units if u.kind == "table_cell")

    # Group units by their owning shape so autofit runs once per shape
    # (not once per paragraph) -- overflow is a property of the box, not
    # of a single paragraph inside it.
    shape_of_unit = {}
    for slide in prs.slides:
        for shape in _flatten(slide.shapes):
            try:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        shape_of_unit[id(para)] = shape
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            for para in cell.text_frame.paragraphs:
                                shape_of_unit[id(para)] = None  # table cells: skip autofit (different mechanism)
            except NotImplementedError:
                # Same unsupported-shape-type issue as in core.py -- skip it here too.
                continue

    for unit in units:
        translated_text = translations.get(id(unit.paragraph))
        if translated_text is None:
            continue
        unit.write(translated_text)

        if autofit and unit.kind == "shape":
            shape = shape_of_unit.get(id(unit.paragraph))
            if shape is not None:
                result = autofit_paragraph(shape, unit.paragraph)
                if result["action"] == "reduced_font_size":
                    report.font_reduced += 1
                elif result["action"] == "reduced_line_spacing":
                    report.line_spacing_reduced += 1
                elif result["action"] == "tightened_char_spacing":
                    report.char_spacing_tightened += 1
                elif result["action"] == "still_overflowing":
                    report.still_overflowing.append(unit.location)

    save_presentation(prs, output_path)
    return report


def _flatten(shapes):
    for shape in shapes:
        try:
            is_group = shape.shape_type == MSO_SHAPE_TYPE.GROUP
        except NotImplementedError:
            is_group = False

        if is_group:
            yield from _flatten(shape.shapes)
        else:
            yield shape