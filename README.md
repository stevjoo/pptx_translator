<div align="center">

# PPTX Translator

**Translate PowerPoint decks (Chinese => English) without breaking the layout.**

Modifies text runs directly inside the `.pptx` file - shapes, images, tables,
positions, animations, and backgrounds are left untouched.

[Features](#features) •
[Quick Start](#quick-start) •
[Usage](#usage) •
[How It Works](#how-it-works) •
[Limitations](#limitations) •
[Roadmap](#roadmap)

</div>

---

## Why this exists

The naive approach to translating a `.pptx` is: extract all text => translate
it => generate a brand new presentation. That almost always breaks the
original layout, because text boxes were sized for the source language, not
the translated one.

This tool takes a different approach: it opens the **existing** `.pptx`,
walks every text run inside it, replaces the run's content with its
translation, and saves the same file back and everything else in the file
(shapes, images, tables, positions, animations, theme, backgrounds) is
byte-for-byte untouched. When a translated string is too long for its box,
a tiered auto-fit engine shrinks it just enough to fit, instead of leaving
it overflowing or silently rebuilding the slide.

## Features

- **Layout-preserving** - edits text runs in place via `python-pptx`, never rebuilds the deck
- **Formatting-aware** - bold/italic/color/font on a run is preserved wherever possible
- **Overflow-aware** - detects text that won't fit and auto-shrinks it (character spacing => font size => line spacing), reporting anything it couldn't fix automatically
- **No budget required** - ships with a free, fully offline translation engine by default (no API key, no sign-up)
- **Pluggable translation backends** - swap in a paid, context-aware LLM engine any time without changing your workflow
- **Covers** - text boxes, titles/placeholders, tables, chart titles, grouped shapes, and speaker notes
- **Resilient** - presentations with SmartArt, OLE objects, or other shape types `python-pptx` doesn't fully support are skipped gracefully instead of crashing the whole run

## Quick Start

```bash
git clone https://github.com/stevjoo/pptx_translator.git
cd pptx_translator

pip install -r requirements.txt
pip install argostranslate          # free, offline translation engine

python -m pptx_translator.cli input.pptx output.pptx
```

That's it - no API key, no cost. The first run downloads the Chinese => English
language model once (needs internet for that step only); every run after
that works fully offline.

## Usage

### Translation backends

| Engine | Cost | Internet required | Quality | Notes |
|---|---|---|---|---|
| `argos` *(default)* | Free | Once, to download the language model | Good | Runs 100% offline afterward |
| `google_free` | Free | Yes, per call | Good for short text | Unofficial endpoint, can rate-limit or break without notice |
| `anthropic` | Paid | Yes | Best, understands full slide context | Needs `ANTHROPIC_API_KEY` |

```bash
# Free, offline (default)
python -m pptx_translator.cli input.pptx output.pptx

# Free, unofficial Google endpoint
pip install deep-translator
python -m pptx_translator.cli input.pptx output.pptx --engine google_free

# Paid, best quality, context-aware across the whole slide
pip install anthropic
export ANTHROPIC_API_KEY=sk-ant-...
python -m pptx_translator.cli input.pptx output.pptx --engine anthropic
```

### Flags

| Flag | Description |
|---|---|
| `--engine {argos,google_free,anthropic}` | Translation backend (default: `argos`) |
| `--model MODEL` | Model name, only used with `--engine anthropic` |
| `--no-notes` | Skip translating speaker notes |
| `--no-autofit` | Disable automatic font-shrink on overflow |

### As a library

```python
from pptx_translator import translate_pptx

report = translate_pptx("input.pptx", "output.pptx", engine="argos")
print(report.as_dict())
```

```python
# Bring your own engine instance for full control (custom model, API key, etc.)
from pptx_translator.translate_engine import AnthropicEngine

report = translate_pptx(
    "input.pptx", "output.pptx",
    engine=AnthropicEngine(model="claude-sonnet-4-6"),
)
```

### Example output

```json
{
  "slides_translated": 28,
  "text_objects": 365,
  "tables_cells": 142,
  "layout_adjustments": {
    "font_size_reduced": 3,
    "line_spacing_reduced": 1,
    "char_spacing_tightened": 2
  },
  "still_overflowing": []
}
```

`still_overflowing` lists any text boxes that could not be made to fit even
after shrinking to the minimum font size and these are worth a manual look.

## How It Works

```
PPTX
 │
 ▼
core.extract_text_units()          => 1 TextUnit per paragraph (shapes, tables, chart titles, notes)
 │
 ▼
translate_engine.translate_units() => batched per slide for context (engine-dependent)
 │
 ▼
TextUnit.write()                   => translation written into run[0]; other runs' formatting kept
 │
 ▼
overflow.autofit_paragraph()       => tiered auto-fit: char spacing → font size → line spacing
 │                                    (bounding-box resize is never automatic, flagged for manual review)
 ▼
core.save_presentation()           => same file, only text runs changed
```

### Coverage

- Text boxes, titles, subtitles, placeholders
- Table cells
- Chart titles (not data labels/axis text)
- Grouped shapes (recursive)
- Speaker notes

## Limitations

1. **Per-run formatting**: if a paragraph mixes formatting mid-sentence
   (e.g. one bold word), the translation is placed entirely into the first
   run, and the other runs are cleared. Safer than mapping runs by index,
   since word order differs between languages but fine-grained formatting
   on that paragraph can be lost.
2. **Overflow detection is an estimate** (font-metric based via Pillow), not
   an actual render. For a final accuracy check, render to images and
   inspect visually:
   ```bash
   soffice --headless --convert-to pdf output.pptx
   pdftoppm -jpeg -r 150 output.pdf slide
   ```
3. **Bounding-box resizing is never automatic** an enlarged box could
   collide with other elements. Anything still overflowing after all
   automatic levels is reported in `still_overflowing` instead.
4. **`argos` and `google_free` translate each paragraph independently**
   they don't use other text on the same slide as context, unlike
   `anthropic`. Fine for factual/technical text; `anthropic` reads more
   naturally when tone/terminology consistency across a slide matters.
5. **Shapes `python-pptx` can't identify** (SmartArt, OLE objects, some
   embedded graphics) are skipped with a `[warn]` message rather than
   crashing the run, text inside them isn't translated and needs manual
   handling in PowerPoint.

## Requirements

- Python 3.9+
- `python-pptx`, `Pillow` (core)
- One of: `argostranslate`, `deep-translator`, or `anthropic` (pick your translation engine)
