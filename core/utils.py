"""
core/utils.py
Shared text-cleaning utilities, fingerprinting, and question de-duplication helpers
used across all parser versions (v1, v2, v3).
"""

from __future__ import annotations

import hashlib
import re
from datetime import date, datetime, time

import pytz

TZ_NAME = "Australia/Sydney"
tz = pytz.timezone(TZ_NAME)


# ---------------------------------------------------------------------------
# Date / time helpers
# ---------------------------------------------------------------------------
def combine_date_time(d: date | None, t: time | None) -> str | None:
    """Combine a date and time into a timezone-aware ISO 8601 string."""
    if not d or not t:
        return None
    dt = tz.localize(datetime.combine(d, t))
    return dt.isoformat()


# ---------------------------------------------------------------------------
# General text helpers (used by v1 / shared code)
# ---------------------------------------------------------------------------
def clean_text(t: str) -> str:
    return re.sub(r"\s+", " ", (t or "")).strip()


def normalize_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


# Q-number prefix pattern (v1)
Q_PREFIX_RE_V1 = re.compile(r"^[^A-Za-z0-9]*(?:lo\s*)?Q\s*\d+\s*[\.\)]\s*", re.IGNORECASE)


def strip_q_prefix(line: str) -> str:
    return Q_PREFIX_RE_V1.sub("", (line or "").strip()).strip()


# ---------------------------------------------------------------------------
# Colour helpers
# ---------------------------------------------------------------------------
def is_red_hex(val: str) -> bool:
    v = (val or "").strip().lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", v):
        return False
    r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
    return r >= 200 and g <= 80 and b <= 80


# ---------------------------------------------------------------------------
# Question fingerprinting & de-duplication (v1 / shared)
# ---------------------------------------------------------------------------
def question_fingerprint(q: dict) -> str:
    qt = normalize_key(q.get("question", ""))
    kind = normalize_key(q.get("kind", ""))
    opts = [normalize_key(x) for x in (q.get("options") or [])]
    pairs = q.get("pairs") or []
    pairs_blob = "||".join(
        [
            normalize_key(p.get("left", "")) + "=>" + normalize_key(p.get("right", ""))
            for p in pairs
        ]
    )
    blob = kind + "||" + qt + "||" + "||".join(opts) + "||" + pairs_blob
    return hashlib.sha1(blob.encode("utf-8")).hexdigest()


def dedupe_questions(questions: list[dict]) -> list[dict]:
    seen: set[str] = set()
    out: list[dict] = []
    for q in questions:
        fp = question_fingerprint(q)
        if fp in seen:
            continue
        seen.add(fp)
        out.append(q)
    return out


# ---------------------------------------------------------------------------
# v2 text helpers
# ---------------------------------------------------------------------------
def v2_clean_text(s: str) -> str:
    s = (s or "").replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def v2_normalize_key(s: str) -> str:
    return v2_clean_text(s).lower()


V2_Q_PREFIX_RE = re.compile(r"^\s*(?:lo\s*)?(?:question|q)\s*\d+\s*[\.\)]\s*", re.IGNORECASE)
V2_ANSWER_GUIDE_INLINE_RE = re.compile(r"\bAnswer\s+(?:may|must|needs)\s+address\b[:\s-]*", re.IGNORECASE)
V2_ANSWER_GUIDE_START_RE = re.compile(r"^\s*Answer\s+(?:may|must|needs)\s+address\b", re.IGNORECASE)
V2_HARD_QNUM_RE = re.compile(r"(?:(?<=\s)|^)(?:lo\s*)?(?:question|q)\s*\d+\s*[\.\)]\s*", re.IGNORECASE)
V2_HARD_QNUM_RANGE_RE = re.compile(r"(?:(?<=\s)|^)(?:lo\s*)?q\s*\d+\s*[-–]\s*\d+\s*[\.\)]\s*", re.IGNORECASE)


def v2_strip_q_prefix(s: str) -> str:
    return v2_clean_text(V2_Q_PREFIX_RE.sub("", s or "", count=1))


def v2_strip_answer_guide(text: str) -> str:
    t = v2_clean_text(text)
    if not t:
        return ""
    m = V2_ANSWER_GUIDE_INLINE_RE.search(t)
    if not m:
        return t
    return v2_clean_text(t[: m.start()])


def v2_trim_after_question_mark(text: str) -> str:
    t = v2_clean_text(text)
    if not t or "?" not in t:
        return t
    return v2_clean_text(t[: t.find("?") + 1])


def v2_split_items_on_internal_qnums(items: list[dict]) -> list[dict]:
    out: list[dict] = []
    for it in items:
        t = v2_clean_text(it.get("text", ""))
        if not t:
            continue
        if V2_HARD_QNUM_RANGE_RE.search(t):
            out.append({"text": t, "is_red": bool(it.get("is_red"))})
            continue
        starts = [m.start() for m in V2_HARD_QNUM_RE.finditer(t)]
        if len(starts) <= 1:
            out.append({"text": t, "is_red": bool(it.get("is_red"))})
            continue
        for a, b in zip(starts, starts[1:] + [len(t)]):
            seg = v2_clean_text(t[a:b])
            if seg:
                out.append({"text": seg, "is_red": bool(it.get("is_red"))})
    return out


# ---------------------------------------------------------------------------
# v3 text helpers
# ---------------------------------------------------------------------------
def v3_clean_text(s: str) -> str:
    s = (s or "").replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def v3_normalize_key(s: str) -> str:
    return v3_clean_text(s).lower()


V3_Q_PREFIX_RE = re.compile(r"^\s*(?:lo\s*)?(?:question|q)\s*\d+\s*[\.\)]\s*", re.IGNORECASE)
V3_ANSWER_GUIDE_INLINE_RE = re.compile(
    r"\bAnswer\s+(?:may|must|need(?:s)?)\s+(?:to\s+)?address\b[:\s-]*", re.IGNORECASE
)
V3_ANSWER_GUIDE_START_RE = re.compile(
    r"^\s*Answer\s+(?:may|must|need(?:s)?)\s+(?:to\s+)?address\b", re.IGNORECASE
)
V3_ANSWER_GUIDE_ANY_RE = re.compile(
    r"\bAnswer\s+(?:may|must|need(?:s)?)\s+(?:to\s+)?address\b", re.IGNORECASE
)
V3_HARD_QNUM_RE = re.compile(
    r"(?:(?<=\s)|^)(?:lo\s*)?(?:question|q)\s*\d+\s*[\.\)]\s*", re.IGNORECASE
)
V3_HARD_QNUM_RANGE_RE = re.compile(
    r"(?:(?<=\s)|^)(?:lo\s*)?q\s*\d+\s*[-–]\s*\d+\s*[\.\)]\s*", re.IGNORECASE
)


def v3_strip_q_prefix(s: str) -> str:
    return v3_clean_text(V3_Q_PREFIX_RE.sub("", s or "", count=1))


def v3_strip_answer_guide(text: str) -> str:
    t = v3_clean_text(text)
    if not t:
        return ""
    m = V3_ANSWER_GUIDE_INLINE_RE.search(t)
    if not m:
        return t
    return v3_clean_text(t[: m.start()])


def v3_trim_after_question_mark(text: str) -> str:
    t = v3_clean_text(text)
    if not t or "?" not in t:
        return t
    return v3_clean_text(t[: t.find("?") + 1])


def v3_trim_after_sentence_if_long(text: str, max_chars: int = 220) -> str:
    t = v3_clean_text(text)
    if len(t) <= max_chars:
        return t
    for sep in [". ", "; ", " - "]:
        pos = t.find(sep)
        if 20 <= pos <= max_chars:
            return v3_clean_text(t[: pos + (1 if sep.startswith(".") else 0)])
    return v3_clean_text(t[:max_chars])


def v3_split_items_on_internal_qnums(items: list[dict]) -> list[dict]:
    out: list[dict] = []
    for it in items:
        t = v3_clean_text(it.get("text", ""))
        if not t:
            continue
        src = it.get("src")
        if V3_HARD_QNUM_RANGE_RE.search(t):
            out.append({"text": t, "is_red": bool(it.get("is_red")), "src": src})
            continue
        starts = [m.start() for m in V3_HARD_QNUM_RE.finditer(t)]
        if len(starts) <= 1:
            out.append({"text": t, "is_red": bool(it.get("is_red")), "src": src})
            continue
        for a, b in zip(starts, starts[1:] + [len(t)]):
            seg = v3_clean_text(t[a:b])
            if seg:
                out.append({"text": seg, "is_red": bool(it.get("is_red")), "src": src})
    return out


# ---------------------------------------------------------------------------
# v3 de-duplication
# ---------------------------------------------------------------------------
def v3_question_dedupe_key(q: dict) -> str:
    kind = (q.get("kind") or "").lower().strip()
    if kind == "matching":
        pairs = q.get("pairs") or []
        parts = [
            f"{v3_normalize_key((p or {}).get('left') or '')}->{v3_normalize_key((p or {}).get('right') or '')}"
            for p in pairs
        ]
        return "matching|" + v3_normalize_key(q.get("question", "")) + "|" + "|".join(parts)
    if kind == "mcq":
        opts = [v3_normalize_key(o) for o in (q.get("options") or []) if v3_normalize_key(o)]
        return "mcq|" + v3_normalize_key(q.get("question", "")) + "|" + "|".join(opts)
    return "essay|" + v3_normalize_key(q.get("question", ""))


def v3_dedupe_questions(questions: list[dict]) -> tuple[list[dict], int]:
    strict_kept: list[dict] = []
    strict_seen: set[str] = set()
    removed = 0
    for q in questions:
        k = v3_question_dedupe_key(q)
        if not k or k in strict_seen:
            removed += 1
            continue
        strict_seen.add(k)
        strict_kept.append(q)

    def text_key(q: dict) -> str:
        return v3_normalize_key(v3_clean_text(q.get("question", "")))

    def kind_rank(q: dict) -> int:
        kind = (q.get("kind") or "").lower().strip()
        if kind == "matching":
            return 3
        if kind == "mcq":
            opts = q.get("options") or []
            return 2 if isinstance(opts, list) and len(opts) >= 2 else 1
        return 0

    best_by_text: dict[str, dict] = {}
    for q in strict_kept:
        tk = text_key(q)
        if not tk:
            continue
        prev = best_by_text.get(tk)
        if prev is None or kind_rank(q) > kind_rank(prev):
            best_by_text[tk] = q

    out: list[dict] = []
    used_ids: set[int] = set()
    for q in strict_kept:
        tk = text_key(q)
        pick = best_by_text.get(tk)
        if pick is None:
            continue
        if id(pick) in used_ids:
            if pick is not q:
                removed += 1
            continue
        if pick is not q:
            removed += 1
            continue
        used_ids.add(id(pick))
        out.append(pick)
    return out, removed


# ---------------------------------------------------------------------------
# v2c (canvasversion2) helpers
# ---------------------------------------------------------------------------
V2C_Q_PREFIX_RE = re.compile(
    r"^[^A-Za-z0-9]*(?:lo\s*)?(?:question|q)\s*\d+\s*(?:[\.\)]\s*|[:\-–—]\s+|\s+)",
    re.IGNORECASE,
)
V2C_NUM_PREFIX_RE = re.compile(r"^[^A-Za-z0-9]*\(?\d+\)?\s*(?:[\.\)]\s*|[:\-–—]\s+)", re.IGNORECASE)


def v2c_clean_text(t: str) -> str:
    return re.sub(r"\s+", " ", (t or "")).strip()


def v2c_normalize_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def v2c_strip_q_prefix(line: str) -> str:
    s = (line or "").strip()
    s = V2C_Q_PREFIX_RE.sub("", s)
    s = V2C_NUM_PREFIX_RE.sub("", s)
    return s.strip()


def v2c_question_fingerprint(q: dict) -> str:
    qt = v2c_normalize_key(q.get("question", ""))
    kind = v2c_normalize_key(q.get("kind", ""))
    opts = [v2c_normalize_key(x) for x in (q.get("options") or [])]
    pairs = q.get("pairs") or []
    pairs_blob = "||".join(
        [
            v2c_normalize_key((p or {}).get("left", "")) + "=>" + v2c_normalize_key((p or {}).get("right", ""))
            for p in pairs
        ]
    )
    blob = kind + "||" + qt + "||" + "||".join(opts) + "||" + pairs_blob
    return hashlib.sha1(blob.encode("utf-8")).hexdigest()


def v2c_dedupe_questions(questions: list[dict]) -> list[dict]:
    seen: set[str] = set()
    out: list[dict] = []
    for q in questions:
        fp = v2c_question_fingerprint(q)
        if fp in seen:
            continue
        seen.add(fp)
        out.append(q)
    return out


def v2c_collapse_duplicate_mcq(questions: list[dict]) -> list[dict]:
    groups: dict[str, list[dict]] = {}
    non_mcq: list[dict] = []

    V2C_NOISE_RE = re.compile(
        r"^(Instructions|For learners|For students|For assessors|Range and conditions|Decision-making rules|"
        r"Pre-approved reasonable adjustments|Rubric|Knowledge Test|"
        r"A rubric has been assigned\b|Answers will be assessed against\b|As a principle\b)\b",
        re.IGNORECASE,
    )
    V2C_OPTION_NOISE_RE = re.compile(
        r"^(Learning\s+Vault|\d{1,2}/\d{1,2}/\d{2,4}|SIT[A-Z0-9]{5,}\b)", re.IGNORECASE
    )

    for q in questions:
        if (q.get("kind") or "").lower() == "mcq":
            key = v2c_normalize_key(q.get("question", ""))
            groups.setdefault(key, []).append(q)
        else:
            non_mcq.append(q)

    def score_mcq(q: dict) -> tuple[int, int]:
        opts = q.get("options") or []
        n = len(opts)
        score = 0
        if 2 <= n <= 6:
            score += 6
        elif 2 <= n <= 10:
            score += 3
        else:
            score -= 3
        if n > 12:
            score -= 8
        bad = sum(
            1
            for o in opts
            if "?" in (o or "")
            or V2C_NOISE_RE.match(o or "")
            or V2C_OPTION_NOISE_RE.match(o or "")
            or v2c_normalize_key(o or "").startswith(("for learners", "for students", "for assessors"))
        )
        score -= bad * 2
        order = int(q.get("_order", 10**9))
        return (score, -order)

    kept: list[dict] = []
    for arr in groups.values():
        arr2 = sorted(arr, key=score_mcq, reverse=True)
        kept.append(arr2[0])

    out = non_mcq + kept
    out.sort(key=lambda q: int(q.get("_order", 10**9)))
    return out