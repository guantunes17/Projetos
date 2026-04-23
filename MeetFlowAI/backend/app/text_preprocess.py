"""
Pré-processamento de transcrições: Unicode, espaços e correção ortográfica leve em português (pyspellchecker).
Controlado por MEETFLOW_PREPROCESS_SPELLCHECK (default: true).
"""

from __future__ import annotations

import os
import re
import unicodedata
from functools import lru_cache

_spellchecker_available = True
try:
    from spellchecker import SpellChecker
except ImportError:
    SpellChecker = None  # type: ignore[misc, assignment]
    _spellchecker_available = False


@lru_cache(maxsize=1)
def _spell_pt() -> SpellChecker | None:
    if not _spellchecker_available or SpellChecker is None:
        return None
    try:
        return SpellChecker(language="pt")
    except Exception:
        return None


def _spell_enabled() -> bool:
    return os.getenv("MEETFLOW_PREPROCESS_SPELLCHECK", "true").lower() in ("1", "true", "yes")


def spell_correct_portuguese(text: str) -> str:
    """Corrige palavras desconhecidas pelo dicionário PT; preserva maiúsculas simples."""
    if not text or not _spell_enabled():
        return text
    spell = _spell_pt()
    if spell is None:
        return text

    word_re = re.compile(r"\w+", re.UNICODE)

    def replace_word(m: re.Match[str]) -> str:
        word = m.group(0)
        lw = word.lower()
        if len(lw) <= 2:
            return word
        if "/" in word or "@" in word or word.startswith("http"):
            return word
        if word.isupper() and len(word) <= 5:
            return word
        if lw.isdigit():
            return word
        if lw in spell:
            return word
        cand = spell.correction(lw)
        if not cand or cand == lw:
            return word
        if word.isupper():
            return cand.upper()
        if word[:1].isupper() and word[1:].islower():
            return cand.capitalize()
        if word[:1].isupper():
            return cand.capitalize()
        return cand

    return word_re.sub(replace_word, text)


def preprocess_transcript(raw: str) -> str:
    """
    NFC, quebras de linha normalizadas, espaços colapsados (comportamento anterior)
    e correção ortográfica opcional antes da API.
    """
    if not raw:
        return ""
    t = unicodedata.normalize("NFC", raw)
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in t.split("\n")]
    t = "\n".join(lines)
    t = re.sub(r"\n{3,}", "\n\n", t)
    t = spell_correct_portuguese(t)
    t = re.sub(r"\s+", " ", t).strip()
    return t
