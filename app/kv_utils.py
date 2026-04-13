"""Shared helpers for cleaning text, normalizing headers, and shaping values."""

from __future__ import annotations

import re
import unicodedata
from decimal import Decimal, InvalidOperation


def clean_text(value) -> str:
    # Convert any cell value to a normalized single-line string for comparisons.
    if value is None:
        return ""
    if isinstance(value, str):
        text = value.strip()
    else:
        text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def normalize_header(value) -> str:
    # Remove accents, spaces, and punctuation so header matching is tolerant.
    text = clean_text(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def flatten_aliases(*alias_maps: dict[str, list[str]]) -> list[str]:
    # Flatten all supported header aliases into one sorted list for error output.
    return sorted(
        {
            alias
            for alias_map in alias_maps
            for aliases in alias_map.values()
            for alias in aliases
        }
    )


def excel_column_letter(index: int) -> str:
    # Convert a 1-based column index to Excel letters: 1 -> A, 27 -> AA.
    if index < 1:
        raise ValueError("Column index must be 1 or greater")
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def slugify(value: str) -> str:
    # Create a simple ASCII code fragment from Vietnamese text.
    normalized = unicodedata.normalize("NFKD", value)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = re.sub(r"[^A-Za-z0-9]+", "-", ascii_text).strip("-").upper()
    return ascii_text or "ITEM"


def make_unique_code(base: str, used: set[str], prefix: str = "") -> str:
    # Reuse the base code when possible, otherwise append -2, -3, ...
    candidate = f"{prefix}{base}"
    if candidate not in used:
        used.add(candidate)
        return candidate
    index = 2
    while True:
        candidate = f"{prefix}{base}-{index}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        index += 1


def to_number(value):
    # Try to preserve numbers as numeric cells; leave unparseable values untouched.
    if value is None or clean_text(value) == "":
        return None
    if isinstance(value, (int, float)):
        return value
    text = clean_text(value).replace(",", "")
    try:
        num = Decimal(text)
    except InvalidOperation:
        return value
    if num == num.to_integral_value():
        return int(num)
    return float(num)


def to_number_or_default(value, default=0):
    # Helper for debt/opening-balance fields where empty means a fixed default.
    number = to_number(value)
    if number is None:
        return default
    return number


def normalize_row(row: list[object], width: int) -> list[object]:
    # Pad or trim template rows so downstream builders can index safely.
    return list(row[:width]) + [""] * max(0, width - len(row))
