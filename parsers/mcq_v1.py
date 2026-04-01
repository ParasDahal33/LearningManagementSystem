"""
parsers/mcq_v1.py
V1 Rule-based parsers for MCQ and Essay questions.
"""

from __future__ import annotations

import re
from core.utils import clean_text, strip_q_prefix

# ---------------------------------------------------------------------------
# Shared noise / question-detection regexes (v1)
# ---------------------------------------------------------------------------
NOISE_RE_V1 = re.compile(
    r"^(Instructions|For learners|For assessors|For students|Range and conditions|Decision-making rules|"
    r"Pre-approved reasonable adjustments|Rubric|Knowledge Test)\b",
    re.IGNORECASE,
)
QUESTION_CMD_INNER_RE_V1 = re.compile(
    r"\b(Which\s+of\s+the\s+following\b|"
    r"(Identify|Select)\s+(one|two|three|four|five|six|seven|eight|nine|ten|\d+)\b)",
    re.IGNORECASE,
)
COMMAND_QUESTION_RE_V1 = re.compile(
    r"^(Illustrate|Critically\s+(?:assess|analyse|analyze|evaluate)|"
    r"Evaluate|Determine|Articulate|Prescribe|Analyse|Analyze|Review|Recommend)\b.+",
    re.IGNORECASE,
)
RUBRIC_START_RE_V1 = re.compile(r"^Answer\s+needs\s+to\s+address\b", re.IGNORECASE)
ESSAY_GUIDE_RE_V1 = re.compile(r"^Answer\s+(may|must)\s+address", re.IGNORECASE)


def parse_mcq_questions_v1(items: list[dict]) -> list[dict]:
    questions_list: list[dict] = []
    current_q: str | None = None
    current_opts: list[dict] = []

    def flush():
        nonlocal current_q, current_opts
        if not current_q:
            return
        opts = [o for o in current_opts if not NOISE_RE_V1.match(o["text"])]
        option_texts = [o["text"] for o in opts]
        correct = [i for i, o in enumerate(opts) if o["is_red"]]
        qtext = strip_q_prefix(current_q.strip())
        qlower = qtext.lower()
        multi = (
            bool(re.search(r"\bselect\s+(two|three|four|five|\d+)", qlower))
            or ("apply" in qlower)
            or (len(correct) > 1)
        )
        questions_list.append(
            {"question": qtext, "options": option_texts, "correct": correct, "multi": multi, "kind": "mcq"}
        )
        current_q = None
        current_opts = []

    for it in items:
        line = clean_text(it.get("text", ""))
        if not line or NOISE_RE_V1.match(line):
            continue
        if ESSAY_GUIDE_RE_V1.match(line):
            current_q = None
            current_opts = []
            continue
        m = QUESTION_CMD_INNER_RE_V1.search(line)
        if m:
            flush()
            start = m.start()
            stem = line[:start].strip()
            cmd_plus = line[start:].strip()
            q_line = f"{stem} {cmd_plus}".strip() if stem else cmd_plus
            current_q = strip_q_prefix(q_line)
            current_opts = []
            continue
        if current_q:
            current_opts.append({"text": line, "is_red": it.get("is_red", False)})

    flush()
    return [
        q
        for q in questions_list
        if len(q.get("options") or []) >= 2 and len(q.get("question") or "") >= 10
    ]


def parse_essay_questions_v1(items: list[dict]) -> list[dict]:
    questions: list[dict] = []
    n = len(items)
    i = 0
    while i < n:
        raw = clean_text(items[i].get("text", ""))
        if not raw or NOISE_RE_V1.match(raw):
            i += 1
            continue
        line = strip_q_prefix(raw)
        if COMMAND_QUESTION_RE_V1.match(line):
            j = i + 1
            next_line = ""
            while j < n:
                nxt = clean_text(items[j].get("text", ""))
                if nxt and not NOISE_RE_V1.match(nxt):
                    next_line = nxt
                    break
                j += 1
            if RUBRIC_START_RE_V1.match(next_line):
                questions.append(
                    {"question": line, "options": [], "correct": [], "multi": False, "kind": "essay"}
                )
                i = j + 1
                continue
        i += 1
    return [q for q in questions if len((q.get("question") or "").strip()) >= 10]