"""Auto-suggest source -> target column mappings using fuzzy matching.

The user always confirms or overrides the result in the UI; this module only
proposes the best guesses so the common case is one click.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from rapidfuzz import fuzz

from .utils import normalize_header

# Below this similarity (0-100) we leave the target unmapped rather than guess.
DEFAULT_THRESHOLD = 72

# ---------------------------------------------------------------------------
# Business-term synonyms — lets the matcher connect headers that mean the same
# thing but don't look alike (e.g. "Client" <-> "Customer Name", "Qty" <->
# "Quantity", "Total" <-> "Amount"). Edit/extend these groups for your domain.
# Each inner list is one group of interchangeable words.
# ---------------------------------------------------------------------------
_SYNONYM_GROUPS: list[list[str]] = [
    ["customer", "client", "buyer", "account", "purchaser", "consignee"],
    ["supplier", "vendor", "seller", "merchant"],
    ["qty", "quantity", "units", "count", "pcs", "pieces", "nos"],
    ["amount", "total", "value", "price", "cost", "subtotal", "sum", "net"],
    ["invoice", "inv", "bill"],
    ["phone", "tel", "telephone", "mobile", "cell", "contact"],
    ["email", "mail", "e mail"],
    ["address", "addr", "location"],
    ["product", "item", "goods", "description", "particulars", "sku"],
    ["date", "dated", "day"],
    ["company", "organisation", "organization", "firm", "business"],
    ["reference", "ref", "po", "order"],
    ["discount", "disc", "rebate"],
    ["tax", "vat", "gst", "duty"],
]

# Tokens that are too generic to match on their own (need a content token too).
_GENERIC_TOKENS = {"number", "no", "num", "id", "name", "date", "code", "type", "ref"}

# token -> canonical group id
_CANON: dict[str, str] = {}
for _gid, _group in enumerate(_SYNONYM_GROUPS):
    for _word in _group:
        _CANON[_word] = f"g{_gid}"


def _canonical_tokens(normalized: str) -> set[str]:
    """Map a normalized header's words to canonical tokens (synonyms unified)."""
    return {_CANON.get(tok, tok) for tok in normalized.split() if tok}


@dataclass
class Suggestion:
    target: str
    source: str | None
    score: float


def _synonym_score(target_tokens: set[str], source_tokens: set[str]) -> float:
    """Score based on shared *meaning*, requiring a non-generic shared token."""
    inter = target_tokens & source_tokens
    if not inter:
        return 0.0
    content = inter - {_CANON.get(t, t) for t in _GENERIC_TOKENS} - _GENERIC_TOKENS
    if not content:
        return 0.0  # only generic tokens overlap — too weak to trust
    jaccard = len(inter) / len(target_tokens | source_tokens)
    return 60.0 + 40.0 * jaccard  # 60..100


def _best_source(target: str, sources: list[str]) -> tuple[str | None, float]:
    norm_target = normalize_header(target)
    tgt_tokens = _canonical_tokens(norm_target)
    best: str | None = None
    best_score = 0.0
    for src in sources:
        norm_src = normalize_header(src)
        if not norm_src or not norm_target:
            continue
        # token_sort handles word reordering; ratio rewards exact-ish matches.
        score = max(
            fuzz.token_sort_ratio(norm_target, norm_src),
            fuzz.ratio(norm_target, norm_src),
        )
        # Strong bonus for containment ("name" in "customer name").
        if norm_target in norm_src or norm_src in norm_target:
            score = max(score, 90.0)
        # Synonym/meaning-based score (handles Qty<->Quantity, Client<->Customer).
        score = max(score, _synonym_score(tgt_tokens, _canonical_tokens(norm_src)))
        if score > best_score:
            best_score = score
            best = src
    return best, best_score


def suggest_mapping(
    target_columns: list[str],
    source_columns: list[str],
    *,
    threshold: float = DEFAULT_THRESHOLD,
) -> dict[str, Suggestion]:
    """Propose a source column for each target column.

    Each source column is used at most once: targets are filled greedily in
    descending score order so the strongest matches win.
    """
    candidates: list[tuple[float, str, str]] = []
    for target in target_columns:
        src, score = _best_source(target, source_columns)
        if src is not None:
            candidates.append((score, target, src))

    candidates.sort(reverse=True, key=lambda t: t[0])
    used_sources: set[str] = set()
    result: dict[str, Suggestion] = {
        t: Suggestion(t, None, 0.0) for t in target_columns
    }

    for score, target, src in candidates:
        if result[target].source is not None:
            continue
        if src in used_sources:
            continue
        if score < threshold:
            continue
        result[target] = Suggestion(target, src, score)
        used_sources.add(src)

    return result


# ---------------------------------------------------------------------------
# Mapping profiles (save / load) so a recurring PDF layout is one click.
# ---------------------------------------------------------------------------

def save_profile(path: str | Path, mapping: dict[str, str | None], meta: dict | None = None) -> None:
    payload = {"meta": meta or {}, "mapping": mapping}
    Path(path).write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_profile(path: str | Path) -> dict[str, str | None]:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    return data.get("mapping", {})
