"""Similaridade textual em Python puro (sem dependencias nativas)."""
from __future__ import annotations

import re
from difflib import SequenceMatcher


def _tokens(text: str) -> list[str]:
    return re.findall(r"\w+", (text or "").lower())


def token_sort_ratio(a: str, b: str) -> float:
    ta = sorted(_tokens(a))
    tb = sorted(_tokens(b))
    sa = " ".join(ta)
    sb = " ".join(tb)
    return SequenceMatcher(None, sa, sb).ratio() * 100.0


def token_set_ratio(a: str, b: str) -> float:
    seta = set(_tokens(a))
    setb = set(_tokens(b))
    if not seta and not setb:
        return 100.0
    inter = seta.intersection(setb)
    only_a = seta - inter
    only_b = setb - inter

    base = " ".join(sorted(inter))
    comb_a = " ".join(sorted(inter.union(only_a)))
    comb_b = " ".join(sorted(inter.union(only_b)))

    r1 = SequenceMatcher(None, base, comb_a).ratio()
    r2 = SequenceMatcher(None, base, comb_b).ratio()
    r3 = SequenceMatcher(None, comb_a, comb_b).ratio()
    return max(r1, r2, r3) * 100.0
