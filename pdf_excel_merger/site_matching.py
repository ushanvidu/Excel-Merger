"""Identify a document's site and match it to a row in the master sheet.

A source document carries a **site key** (Site ID or Site Name). The master
sheet has one row per site, keyed by a column the user designates. Matching is:

  1. **Exact** on a normalized key (``"COL-001"`` == ``"col 001"`` == ``"COL001"``)
     — the reliable path, ideal for site codes/IDs.
  2. **Fuzzy** fallback on the original text (for site *names* with small
     spelling/format differences), gated by a high threshold and reported so the
     user can see it wasn't an exact hit.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from rapidfuzz import fuzz

from .utils import clean_cell, normalize_header

# Header words that suggest a column holds a site identifier.
_KEY_HINTS = [
    "site id", "siteid", "site code", "sitecode", "site name", "sitename",
    "site no", "site number", "site", "cell id", "cellid", "node", "node id",
    "enodeb", "enb", "bts", "nodeb", "gnb", "tower id", "tower", "station id",
    "station", "location id", "location code", "id", "code",
]
_KEY_HINTS_NORM = [normalize_header(h) for h in _KEY_HINTS]

DEFAULT_FUZZY_THRESHOLD = 88


def normalize_key(value) -> str:
    """Aggressively normalize a site key for exact comparison.

    Uppercases and strips every non-alphanumeric character, so separators and
    spacing differences don't break a match. Leading zeros are kept (``COL001``
    and ``COL01`` are genuinely different sites).
    """
    return re.sub(r"[^A-Z0-9]", "", clean_cell(value).upper())


def score_key_column(column_name: str) -> float:
    """How strongly a column header looks like a site-key column (0-100)."""
    norm = normalize_header(column_name)
    if not norm:
        return 0.0
    best = 0.0
    for hint in _KEY_HINTS_NORM:
        if norm == hint:
            return 100.0
        if hint in norm or norm in hint:
            best = max(best, 88.0)
        best = max(best, float(fuzz.token_sort_ratio(norm, hint)))
    # A column literally containing "site" is a strong signal.
    if "site" in norm:
        best = max(best, 90.0)
    return best


def detect_key_column(columns: list[str]) -> str | None:
    """Pick the most likely site-key column from a list of headers."""
    ranked = sorted(columns, key=score_key_column, reverse=True)
    if ranked and score_key_column(ranked[0]) >= 60:
        return ranked[0]
    return None


@dataclass
class MatchResult:
    row: int | None        # matched master-sheet row number (None if unmatched)
    matched_value: str     # the master key text that was matched ("" if none)
    method: str            # "exact" | "fuzzy" | "unmatched" | "empty"
    score: float


class SiteIndex:
    """Index of master-sheet site keys for fast/fuzzy lookup.

    ``entries`` is a list of ``(row_number, original_key_text)``.
    """

    def __init__(self, entries: list[tuple[int, str]]):
        self.entries = entries
        self._exact: dict[str, tuple[int, str]] = {}
        for row, original in entries:
            nk = normalize_key(original)
            if nk and nk not in self._exact:  # first occurrence wins
                self._exact[nk] = (row, original)

    def match(
        self,
        raw_key,
        *,
        allow_fuzzy: bool = True,
        threshold: float = DEFAULT_FUZZY_THRESHOLD,
    ) -> MatchResult:
        nk = normalize_key(raw_key)
        if not nk:
            return MatchResult(None, "", "empty", 0.0)

        hit = self._exact.get(nk)
        if hit is not None:
            return MatchResult(hit[0], hit[1], "exact", 100.0)

        if allow_fuzzy:
            raw_norm = normalize_header(raw_key)
            best_row: int | None = None
            best_val = ""
            best_score = 0.0
            for row, original in self.entries:
                score = max(
                    fuzz.token_sort_ratio(raw_norm, normalize_header(original)),
                    fuzz.ratio(nk, normalize_key(original)),
                )
                if score > best_score:
                    best_score, best_row, best_val = score, row, original
            if best_row is not None and best_score >= threshold:
                return MatchResult(best_row, best_val, "fuzzy", best_score)

        return MatchResult(None, "", "unmatched", 0.0)
