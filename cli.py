"""
cli.py

Usage (free, offline, no API key -- default):
    pip install argostranslate
    python -m pptx_translator.cli input.pptx output.pptx

Usage (free, unofficial Google endpoint):
    pip install deep-translator
    python -m pptx_translator.cli input.pptx output.pptx --engine google_free

Usage (paid, best quality, context-aware):
    pip install anthropic
    export ANTHROPIC_API_KEY=sk-ant-...
    python -m pptx_translator.cli input.pptx output.pptx --engine anthropic
"""

import argparse
import json
import sys

from .translator import translate_pptx


def main():
    parser = argparse.ArgumentParser(
        description="Translate a .pptx from Chinese to English while preserving layout."
    )
    parser.add_argument("input", help="Path to the input .pptx file")
    parser.add_argument("output", help="Path to write the translated .pptx file")
    parser.add_argument(
        "--engine",
        default="argos",
        choices=["anthropic", "argos", "google_free"],
        help="Translation backend. 'argos' (default) is free and offline. "
             "'google_free' is free but uses an unofficial endpoint. "
             "'anthropic' is paid but the highest quality and slide-aware.",
    )
    parser.add_argument(
        "--model",
        default="claude-sonnet-4-6",
        help="Model name, only used when --engine anthropic",
    )
    parser.add_argument("--no-notes", action="store_true", help="Skip translating speaker notes")
    parser.add_argument("--no-autofit", action="store_true", help="Disable auto font-shrink on overflow")
    args = parser.parse_args()

    engine_kwargs = {"model": args.model} if args.engine == "anthropic" else {}

    report = translate_pptx(
        input_path=args.input,
        output_path=args.output,
        engine=args.engine,
        engine_kwargs=engine_kwargs,
        include_notes=not args.no_notes,
        autofit=not args.no_autofit,
    )

    print(f"\n✅ Translation complete -> {args.output}\n")
    print(json.dumps(report.as_dict(), indent=2, ensure_ascii=False))

    if report.still_overflowing:
        print(
            f"\n⚠️  {len(report.still_overflowing)} text box(es) may still overflow "
            "even after shrinking the font to its minimum. Review manually or "
            "render to images to check visually:",
            file=sys.stderr,
        )
        for loc in report.still_overflowing:
            print(f"   - {loc}", file=sys.stderr)


if __name__ == "__main__":
    main()
