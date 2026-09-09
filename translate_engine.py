"""
translate_engine.py
Pluggable translation backends. Pick whichever fits your budget:

    - AnthropicEngine   : paid (Claude API key required). Best quality,
                          understands per-slide context, follows the
                          terminology/format rules in SYSTEM_PROMPT.
    - ArgosEngine       : FREE, runs fully offline/locally after a one-time
                          model download. No API key. Translates each string
                          independently (no cross-item slide context).
    - GoogleFreeEngine  : FREE, no API key, uses the unofficial public
                          Google Translate web endpoint via `deep-translator`.
                          Needs internet per call. Can break or get
                          rate-limited without warning -- fine for a
                          prototype, not something to depend on in production.

All engines implement the same interface: translate_batch(list[str]) -> list[str]
so translate_units() and the rest of the pipeline don't care which one you use.
"""

from __future__ import annotations

import json
from itertools import groupby
from typing import Optional

from .core import TextUnit

SYSTEM_PROMPT = """You are a professional presentation translator.
Translate the given Chinese (Mandarin) slide text into natural,
academic-register English suitable for a formal presentation.

Rules:
- Preserve meaning exactly. Do not add or remove information.
- Keep technical terms, abbreviations, numbers, citations, and proper nouns
  unchanged unless there is a standard English equivalent.
- Keep each item's line breaks ("\\n") in roughly the same place if present.
- Do not translate item-by-item in isolation -- use the other items on the
  same slide as context for consistent terminology.
- Return ONLY a JSON array of strings, same length and same order as the
  input array. No markdown fences, no commentary, no extra keys.
"""


class Engine:
    """Base interface every translation backend must implement."""

    def translate_batch(self, texts: list[str]) -> list[str]:
        raise NotImplementedError

    def supports_context(self) -> bool:
        """
        Whether this engine can use other items on the same slide as
        context (see point #7 in the original discussion: translating
        objects one by one, out of context, tends to read unnaturally).
        """
        return False


# ---------------------------------------------------------------------------
# Paid option: Anthropic API (best quality, context-aware)
# ---------------------------------------------------------------------------

class AnthropicEngine(Engine):
    """
    Requires: pip install anthropic
    Requires: export ANTHROPIC_API_KEY=sk-ant-...
    """

    def __init__(self, model: str = "claude-sonnet-4-6", client=None):
        import anthropic  # imported lazily so this dependency is optional
        self.model = model
        self.client = client or anthropic.Anthropic()

    def supports_context(self) -> bool:
        return True

    def translate_batch(self, texts: list[str]) -> list[str]:
        if not texts:
            return []
        payload = json.dumps(texts, ensure_ascii=False)
        message = self.client.messages.create(
            model=self.model,
            max_tokens=4000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": f"Input JSON array:\n{payload}"}],
        )
        raw = "".join(b.text for b in message.content if b.type == "text").strip()
        raw = raw.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, list) and all(isinstance(x, str) for x in parsed):
                return parsed
        except json.JSONDecodeError:
            pass
        # Fallback: translate one item at a time so a single malformed
        # response doesn't drop the whole slide's text.
        return [self._translate_one(t) for t in texts]

    def _translate_one(self, text: str) -> str:
        result = self.translate_batch([text]) if text else [text]
        return result[0] if result else text


# ---------------------------------------------------------------------------
# Free option #1: Argos Translate (fully offline, no API key, no per-call cost)
# ---------------------------------------------------------------------------

class ArgosEngine(Engine):
    """
    Requires: pip install argostranslate
    First run needs internet ONCE to download the zh->en language package
    (a few dozen MB). After that it runs 100% locally/offline, no API key,
    no per-request cost, no rate limit.

    Trade-off: translates each string independently -- it does not use
    other text on the same slide as context, so terminology consistency
    across a slide relies purely on the model itself, not on prompting.
    """

    def __init__(self, from_code: str = "zh", to_code: str = "en", auto_install: bool = True):
        import argostranslate.package
        import argostranslate.translate

        self._pkg = argostranslate.package
        self._translate_mod = argostranslate.translate
        self.from_code = from_code
        self.to_code = to_code

        if auto_install:
            self._ensure_package_installed()

    def _ensure_package_installed(self) -> None:
        installed = self._translate_mod.get_installed_languages()
        has_pair = any(
            lang.code == self.from_code and any(t.to_lang.code == self.to_code for t in lang.translations_from)
            for lang in installed
        )
        if has_pair:
            return

        print(f"[argos] '{self.from_code}->{self.to_code}' package not installed yet. "
              f"Downloading once (needs internet for this step only)...")
        self._pkg.update_package_index()
        available = self._pkg.get_available_packages()
        match = next(
            (p for p in available if p.from_code == self.from_code and p.to_code == self.to_code),
            None,
        )
        if match is None:
            raise RuntimeError(
                f"No Argos package found for {self.from_code}->{self.to_code}. "
                "Check available languages at https://www.argosopentech.com/argospm/index/"
            )
        download_path = match.download()
        self._pkg.install_from_path(download_path)
        print("[argos] Package installed. Future runs will work fully offline.")

    def translate_batch(self, texts: list[str]) -> list[str]:
        out = []
        for t in texts:
            if not t.strip():
                out.append(t)
                continue
            out.append(self._translate_mod.translate(t, self.from_code, self.to_code))
        return out


# ---------------------------------------------------------------------------
# Free option #2: unofficial Google Translate endpoint (via deep-translator)
# ---------------------------------------------------------------------------

class GoogleFreeEngine(Engine):
    """
    Requires: pip install deep-translator
    No API key needed. Uses Google Translate's public web endpoint, which is
    NOT an official/supported API -- it can rate-limit or break without
    notice. Good for trying things out for free; do not rely on it for
    anything that must always work.

    source: "zh-CN" for Simplified Chinese (default), "zh-TW" for
    Traditional Chinese -- use whichever matches your slides.
    """

    def __init__(self, source: str = "zh-CN", target: str = "en"):
        from deep_translator import GoogleTranslator
        self.source = source
        self.target = target
        self._cls = GoogleTranslator

    def translate_batch(self, texts: list[str]) -> list[str]:
        translator = self._cls(source=self.source, target=self.target)
        out = []
        for t in texts:
            if not t.strip():
                out.append(t)
                continue
            out.append(translator.translate(t))
        return out


# ---------------------------------------------------------------------------
# Engine factory + orchestration
# ---------------------------------------------------------------------------

ENGINES = {
    "anthropic": AnthropicEngine,
    "argos": ArgosEngine,
    "google_free": GoogleFreeEngine,
}


def build_engine(name: str, **kwargs) -> Engine:
    if name not in ENGINES:
        raise ValueError(f"Unknown engine '{name}'. Choose from: {list(ENGINES)}")
    return ENGINES[name](**kwargs)


def translate_units(units: list[TextUnit], engine: Engine) -> dict[int, str]:
    """
    Translate all units, grouped by slide_index so context-aware engines
    (like AnthropicEngine) can use other text on the same slide as context.
    Engines that don't support context just translate each string on its own
    regardless of grouping -- grouping is harmless either way.

    Returns: {id(unit.paragraph): translated_text}, used by the caller to
    write results back one by one via TextUnit.write().
    """
    results: dict[int, str] = {}

    units_sorted = sorted(units, key=lambda u: u.slide_index)
    for _slide_idx, group_iter in groupby(units_sorted, key=lambda u: u.slide_index):
        group = list(group_iter)
        texts = [u.original_text for u in group]
        translated = engine.translate_batch(texts)

        if len(translated) != len(texts):
            # Safety net: if the batch call returned a mismatched count,
            # fall back to translating one item at a time so nothing is
            # silently dropped or misaligned.
            translated = [engine.translate_batch([t])[0] if t.strip() else t for t in texts]

        for unit, t in zip(group, translated):
            results[id(unit.paragraph)] = t

    return results
