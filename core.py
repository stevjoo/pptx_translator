"""
core.py
Extracts and writes back text in a .pptx file, working at the RUN level
(not shape.text / text_frame.text) so bold/italic/color/font per word is
not lost.

Concept:
    Presentation
     └── Slide
          └── Shape (can be nested: a GroupShape contains other shapes)
               ├── TextFrame
               │    └── Paragraph
               │         └── Run   <-- smallest unit with independent formatting
               └── Table
                    └── Cell (has its own text_frame, same as a regular shape)

We never touch: position, size, images, background, animation, chart XML,
or SmartArt. Only the content of text runs is replaced.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterator, Optional

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.text.text import _Paragraph, _Run
from pptx.util import Pt


@dataclass
class TextUnit:
    """
    A single "text unit" that gets translated as one piece. This is usually
    one PARAGRAPH (all its runs joined together), because a sentence that's
    split run-by-run and translated one fragment at a time tends to lose
    context (see point #7 in the original design discussion).

    runs        : direct references to the python-pptx run objects, used to
                  write the translated result back without breaking formatting.
    slide_index : used to group text by slide for context-aware translation.
    kind        : "shape" | "table_cell" | "notes" | "chart_title"
    location    : human-readable info for logging/debugging, e.g.
                  "Slide 3 > Table 1 > R2C1"
    """
    slide_index: int
    kind: str
    location: str
    paragraph: _Paragraph
    original_text: str = field(init=False)

    def __post_init__(self):
        self.original_text = self.paragraph.text

    # ---- writing the translation back ----
    def write(self, translated_text: str) -> None:
        """
        Write the translated text into this paragraph WITHOUT deleting the
        other run objects entirely, when possible.

        Strategy:
        - If the paragraph has only 1 run -> safest case, just overwrite
          run.text directly.
        - If the paragraph has >1 run (mixed formatting, e.g. one bold
          word) -> put the ENTIRE new text into the FIRST run (inheriting
          its formatting), then clear the remaining runs. This is a
          deliberate trade-off: better to lose fine-grained formatting on
          that one paragraph than to assign run.text by index, since word
          order differs between languages -- the 2nd run in Chinese is
          not necessarily the 2nd run in English.
        """
        runs = self.paragraph.runs
        if not runs:
            # Paragraph has no runs at all (rare) -> create a new run
            self.paragraph.text = translated_text
            return

        runs[0].text = translated_text
        for r in runs[1:]:
            r.text = ""

    def has_bold_anywhere(self) -> bool:
        return any(r.font.bold for r in self.paragraph.runs)


def _safe_shape_type(shape) -> str:
    """shape.shape_type raises NotImplementedError for some SmartArt/OLE/
    embedded-object shapes. Fall back to a generic label instead of crashing."""
    try:
        return str(shape.shape_type)
    except NotImplementedError:
        return "UNKNOWN_SHAPE_TYPE"


def _iter_shapes(shapes) -> Iterator:
    """Recursively yield shapes, descending into GroupShape.

    Some shapes (SmartArt, OLE objects, certain embedded graphics) raise
    NotImplementedError when python-pptx tries to determine their
    shape_type. We treat those as opaque leaf shapes -- we can't descend
    into them, but we still yield them so has_text_frame/has_table/etc.
    checks downstream get a chance to run (most just return False safely).
    """
    for shape in shapes:
        try:
            is_group = shape.shape_type == MSO_SHAPE_TYPE.GROUP
        except NotImplementedError:
            is_group = False

        if is_group:
            yield from _iter_shapes(shape.shapes)
        else:
            yield shape


def _collect_from_shape(units: list[TextUnit], shape, s_idx: int) -> None:
    """Extract TextUnits from a single shape (text frame, table, or chart title)."""
    # --- regular text frame (title, textbox, placeholder) ---
    if shape.has_text_frame:
        shape_type_label = _safe_shape_type(shape)
        for p_idx, para in enumerate(shape.text_frame.paragraphs):
            if para.text.strip():
                units.append(TextUnit(
                    slide_index=s_idx,
                    kind="shape",
                    location=f"Slide {s_idx + 1} > {shape_type_label} '{shape.name}' > P{p_idx}",
                    paragraph=para,
                ))

    # --- table ---
    if shape.has_table:
        table = shape.table
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                for p_idx, para in enumerate(cell.text_frame.paragraphs):
                    if para.text.strip():
                        units.append(TextUnit(
                            slide_index=s_idx,
                            kind="table_cell",
                            location=f"Slide {s_idx + 1} > Table '{shape.name}' > R{r_idx}C{c_idx} > P{p_idx}",
                            paragraph=para,
                        ))

    # --- chart title: python-pptx accesses this differently, so it is
    #     handled separately to keep the caller's loop simple ---
    if shape.has_chart:
        chart = shape.chart
        if chart.has_title and chart.chart_title.text_frame.text.strip():
            for p_idx, para in enumerate(chart.chart_title.text_frame.paragraphs):
                if para.text.strip():
                    units.append(TextUnit(
                        slide_index=s_idx,
                        kind="chart_title",
                        location=f"Slide {s_idx + 1} > Chart '{shape.name}' > Title",
                        paragraph=para,
                    ))


def extract_text_units(
    prs: Presentation,
    include_notes: bool = True,
    slide_indices: Optional[list[int]] = None,
) -> list[TextUnit]:
    """
    Collect every TextUnit (1 per paragraph) that has non-empty text, from
    all slides (or a subset given by slide_indices).
    """
    units: list[TextUnit] = []

    for s_idx, slide in enumerate(prs.slides):
        if slide_indices is not None and s_idx not in slide_indices:
            continue

        for shape in _iter_shapes(slide.shapes):
            try:
                _collect_from_shape(units, shape, s_idx)
            except NotImplementedError:
                # Some shape types (SmartArt, OLE objects, certain embedded
                # graphics) aren't fully supported by python-pptx. Skip that
                # one shape rather than aborting the whole presentation.
                print(f"[warn] Slide {s_idx + 1}: skipped an unsupported shape ('{getattr(shape, 'name', '?')}')")
                continue

        # --- speaker notes ---
        if include_notes and slide.has_notes_slide:
            notes_tf = slide.notes_slide.notes_text_frame
            for p_idx, para in enumerate(notes_tf.paragraphs):
                if para.text.strip():
                    units.append(TextUnit(
                        slide_index=s_idx,
                        kind="notes",
                        location=f"Slide {s_idx + 1} > Notes > P{p_idx}",
                        paragraph=para,
                    ))

    return units


def open_presentation(path: str) -> Presentation:
    return Presentation(path)


def save_presentation(prs: Presentation, path: str) -> None:
    prs.save(path)