"""
parsers/mcq_v3.py
V3 parsers: rule-based fallback and AI filtering logic.
"""

from __future__ import annotations

import re
from core.utils import (
    v3_clean_text,
    v3_normalize_key,
    v3_strip_q_prefix,
    v3_strip_answer_guide,
    v3_trim_after_question_mark,
    v3_trim_after_sentence_if_long,
    V3_ANSWER_GUIDE_START_RE,
    V3_ANSWER_GUIDE_ANY_RE,
)


V3_IGNORE_LINE_RE = re.compile(
    r"^(instructions|for learners|for assessors|for students|range and conditions|decision-making rules|pre-approved|rubric|feedback|knowledge test)\b",
    re.IGNORECASE,
)
V3_IGNORE_SECTION_RE = re.compile(
    r"^(?:range and conditions|decision-making rules|pre-approved reasonable adjustments|rubric)\b",
    re.IGNORECASE,
)
V3_IGNORE_TABLE_RE = re.compile(
    r"^(?:poultry ingredient|definition|style/method of cooking|poultry type or cut|essential characteristics|classical chicken dishes|contemporary chicken dishes)\b",
    re.IGNORECASE,
)
V3_COOKERY_METHOD_WORD_RE = re.compile(
    r"^(?:pan[-\s]?fry|deep[-\s]?fry|stir[-\s]?fry|roast|bake|grill|bbq|braise|stew|simmer|poach|saute|sauté|steam|boil)\b",
    re.IGNORECASE,
)
V3_QUESTION_START_RE = re.compile(
    r"^(?:(?:lo\s*)?(?:question|q)\s*\d+\s*[\.\)]\s*)?"
    r"(?:critically\s+)?"
    r"(?:which of the following|select|choose|pick|match|complete|list|name|identify|define|describe|explain|"
    r"outline|state|provide|illustrate|evaluate|determine|articulate|discuss|analyse|analyze|compare|review|"
    r"appraise|assess|what|when|where|why|how|must\b)\b",
    re.IGNORECASE,
)
V3_OPTION_LINE_RE = re.compile(
    r"^\s*(?:(?:option\s*\d+)|(?:\(?[a-h]\)|[a-h][\.\)])|(?:\(?i{1,3}v?\)|i{1,3}v?[\.\)]))\s+",
    re.IGNORECASE,
)


def _v3_looks_like_question_start(text: str) -> bool:
    t = v3_clean_text(text)
    if not t:
        return False
    if t.endswith("?") or "____" in t or "___" in t:
        return True
    return bool(V3_QUESTION_START_RE.match(t))


def _v3_looks_like_answer_guide_bullet(text: str) -> bool:
    t = v3_clean_text(text)
    if not t:
        return False
    tl = t.lower()
    return (
        tl.startswith(("answer may address", "answer must address", "answer needs to address"))
        or tl in {"that is blank", "has nothing written in the space provided"}
        or tl.startswith("does not attempt to answer")
    )


def _v3_looks_like_option_line(text: str) -> bool:
    t = v3_clean_text(text)
    if not t:
        return False
    return bool(V3_OPTION_LINE_RE.match(t)) or t.lower().startswith(("true", "false"))


def _v3_is_admin_or_meta_line(text: str) -> bool:
    t = v3_clean_text(text)
    if not t:
        return True
    tl = t.lower()
    if tl.startswith("when you have completed all questions") or tl.startswith("by submitting your") or tl.startswith("where a learner is assessed as"):
        return True
    if V3_IGNORE_LINE_RE.match(t) or V3_IGNORE_SECTION_RE.match(t) or V3_IGNORE_TABLE_RE.match(t):
        return True
    if _v3_looks_like_answer_guide_bullet(t) or V3_ANSWER_GUIDE_START_RE.match(t):
        return True
    return False


def v3_parse_essay_questions_rule_based(items: list[dict]) -> list[dict]:
    out: list[dict] = []
    seen: set[str] = set()
    in_answer_guide = False
    mcqish_stem_re = re.compile(r"^\s*(select|choose|pick|match|complete)\b", re.IGNORECASE)

    def has_answer_guide_soon(idx: int) -> bool:
        for j in range(idx + 1, min(len(items), idx + 8)):
            t2 = v3_clean_text(items[j].get("text", ""))
            if not t2:
                continue
            if V3_ANSWER_GUIDE_START_RE.match(t2) or V3_ANSWER_GUIDE_ANY_RE.search(t2):
                return True
        return False

    i = 0
    while i < len(items):
        t = v3_clean_text(items[i].get("text", ""))
        if not t or _v3_is_admin_or_meta_line(t) or bool(items[i].get("is_red")):
            i += 1
            continue
        if V3_ANSWER_GUIDE_START_RE.match(t) or V3_ANSWER_GUIDE_ANY_RE.search(t):
            in_answer_guide = True
            i += 1
            continue
        if V3_ANSWER_GUIDE_ANY_RE.search(t):
            t = v3_strip_answer_guide(t)
        stem = v3_trim_after_question_mark(v3_strip_q_prefix(t))
        stem = v3_trim_after_sentence_if_long(stem)
        if not stem or len(stem) < 10 or not _v3_looks_like_question_start(stem):
            i += 1
            continue
        if "which of the following" in stem.lower() or mcqish_stem_re.match(stem):
            i += 1
            continue
        if in_answer_guide and not (stem.endswith("?") or has_answer_guide_soon(i)):
            i += 1
            continue

        optionish = sum(
            1
            for j in range(i + 1, min(len(items), i + 8))
            if v3_clean_text(items[j].get("text", ""))
            and not _v3_is_admin_or_meta_line(v3_clean_text(items[j].get("text", "")))
            and _v3_looks_like_option_line(v3_clean_text(items[j].get("text", "")))
        )
        if optionish >= 2:
            i += 1
            continue

        if not stem.endswith("?") and not has_answer_guide_soon(i):
            i += 1
            continue

        k = v3_normalize_key(stem)
        if k and k not in seen:
            seen.add(k)
            out.append({
                "question": stem,
                "options": [],
                "correct": [],
                "multi": False,
                "kind": "essay",
                "_order": i,
                "qnum": None,
            })

        in_answer_guide = has_answer_guide_soon(i)
        i += 1
    return out


def v3_filter_items_for_ai(
    items: list[dict],
    ignore_terms: set[str] | None = None,
    ignore_texts: set[str] | None = None,
    mode: str = "balanced",
) -> list[dict]:
    out: list[dict] = []
    ignore_terms = ignore_terms or set()
    ignore_terms_norm = {v3_normalize_key(t) for t in ignore_terms if v3_normalize_key(t)}
    ignore_term_prefixes = sorted(ignore_terms_norm, key=len, reverse=True)
    ignore_texts_norm = {v3_normalize_key(t) for t in (ignore_texts or set()) if v3_normalize_key(t)}
    mode = (mode or "balanced").strip().lower()
    if mode not in {"balanced", "loose", "strict"}:
        mode = "balanced"
    in_answer_guide = False

    for it in items:
        t = v3_clean_text(it.get("text", ""))
        if not t:
            continue
        if v3_normalize_key(t) in ignore_texts_norm:
            continue
        if V3_IGNORE_LINE_RE.match(t) or V3_IGNORE_SECTION_RE.match(t):
            continue
        if mode in {"balanced", "strict"}:
            if V3_ANSWER_GUIDE_START_RE.match(t):
                in_answer_guide = True
                continue
            if V3_ANSWER_GUIDE_ANY_RE.search(t):
                pre = v3_strip_answer_guide(t)
                if pre and _v3_looks_like_question_start(pre) and len(pre) >= 10:
                    out.append({"text": pre, "is_red": False})
                in_answer_guide = True
                continue
        if V3_IGNORE_TABLE_RE.match(t):
            continue
        if mode in {"balanced", "strict"} and _v3_looks_like_answer_guide_bullet(t):
            continue
        tn = v3_normalize_key(t)
        if tn in ignore_terms_norm:
            continue
        if any(tn.startswith(pref + " ") for pref in ignore_term_prefixes):
            continue
        m = re.match(r"^\s*(?:q\s*\d+\s*[\.\)]\s*)?(.*)$", t, flags=re.IGNORECASE)
        if m and v3_normalize_key(v3_clean_text(m.group(1))) in ignore_terms_norm:
            continue
        if mode == "strict" and in_answer_guide:
            if _v3_looks_like_question_start(t):
                in_answer_guide = False
            else:
                continue
        if mode == "balanced" and in_answer_guide:
            if _v3_looks_like_question_start(t):
                in_answer_guide = False
            else:
                if len(t) <= 80 and not t.endswith("?"):
                    continue
                if bool(it.get("is_red")) and len(t) <= 80:
                    continue
                continue
        if mode in {"balanced", "strict"} and len(t) <= 20 and V3_COOKERY_METHOD_WORD_RE.match(t):
            continue
        out.append(it)
    return out