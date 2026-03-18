"""
parsers/mcq_parser.py
Multiple-choice and essay question parsing for all parser versions (v1, v2c, v3).
"""

from __future__ import annotations

import re

from .util import (
    clean_text,
    normalize_key,
    strip_q_prefix,
    v2c_clean_text,
    v2c_normalize_key,
    v2c_strip_q_prefix,
    v3_clean_text,
    v3_normalize_key,
    v3_strip_q_prefix,
    v3_strip_answer_guide,
    v3_trim_after_question_mark,
    v3_trim_after_sentence_if_long,
    V3_ANSWER_GUIDE_START_RE,
    V3_ANSWER_GUIDE_ANY_RE,
)

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


# ===========================================================================
# V1 Parsers
# ===========================================================================
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


# ===========================================================================
# V2C (canvasversion2) Parsers
# ===========================================================================
V2C_NOISE_RE = re.compile(
    r"^(Instructions|For learners|For students|For assessors|Range and conditions|Decision-making rules|"
    r"Pre-approved reasonable adjustments|Rubric|Knowledge Test|"
    r"A rubric has been assigned\b|Answers will be assessed against\b|As a principle\b)\b",
    re.IGNORECASE,
)
V2C_STOP_OPTION_RE = re.compile(
    r"^(Learner feedback|Assessment outcome|Assessor signature|Assessor name|Final comments)\b"
    r"|^(Competent|Not Yet Competent|NYC|C|Date)\s*[:\-]?$",
    re.IGNORECASE,
)
V2C_OPTION_NOISE_RE = re.compile(
    r"^(Learning\s+Vault|\d{1,2}/\d{1,2}/\d{2,4}|SIT[A-Z0-9]{5,}\b)", re.IGNORECASE
)
V2C_QUESTION_CMD_INNER_RE = re.compile(
    r"\b(Which\s+of\s+the\s+following\b|"
    r"(Identify|Select|Choose|Pick)\s+(?:the\s+)?(one|two|three|four|five|six|seven|eight|nine|ten|\d+)\b)",
    re.IGNORECASE,
)
V2C_COMMAND_QUESTION_RE = re.compile(
    r"^(Illustrate|Explain|Describe|Discuss|Outline|Compare|Summari[sz]e|"
    r"Critically\s+(?:assess|analyse|analyze|evaluate)|"
    r"Evaluate|Determine|Articulate|Prescribe|Analyse|Analyze|Review|Recommend|Provide)\b.+",
    re.IGNORECASE,
)
V2C_RUBRIC_START_RE = re.compile(r"^Answer\s+needs\s+to\s+address\b", re.IGNORECASE)
V2C_ESSAY_GUIDE_RE = re.compile(r"^Answer\s+(may|must)\s+address", re.IGNORECASE)
V2C_MATCHING_STEM_RE = re.compile(
    r"\b(complete\s+the\s+table|drag(?:ging)?\s+and\s+drop(?:ping)?|drag\s+and\s+drop|"
    r"match\s+each|match\s+the\s+following|match\s+.*\s+to\s+the\s+correct|select\s+one.*for\s+each)\b",
    re.IGNORECASE,
)
V2C_LETTERED_OPT_PREFIX_RE = re.compile(r"^\s*(?:[\(\[]?[a-hA-H][\)\].:-])\s+")
V2C_DANGLING_Q_END_RE = re.compile(r"\b(of|for|to|with|and|or|in|on|at|from|by|as|about)\s*$", re.IGNORECASE)


def _v2c_looks_like_matching_stem(t: str) -> bool:
    t2 = v2c_strip_q_prefix(v2c_clean_text(t))
    if not t2:
        return False
    low = t2.lower()
    if low.startswith(("for learners", "for assessors", "for students")):
        return False
    if V2C_COMMAND_QUESTION_RE.match(t2):
        return False
    if "which of the following" in low:
        return False
    return bool(V2C_MATCHING_STEM_RE.search(t2))


def v2c_merge_dangling_question_lines(items: list[dict]) -> list[dict]:
    out: list[dict] = []
    i = 0
    n = len(items)
    while i < n:
        it = items[i]
        t = v2c_clean_text(it.get("text", ""))
        if not t:
            i += 1
            continue
        t_stem = v2c_strip_q_prefix(t)
        can_start_q = bool(re.match(r"^(?:in\s+)?(which|what|why|how)\b", t_stem, re.IGNORECASE))
        dangling = (
            can_start_q
            and "?" not in t_stem
            and V2C_DANGLING_Q_END_RE.search(t_stem)
            and not _v2c_looks_like_matching_stem(t_stem)
            and not V2C_NOISE_RE.match(t_stem)
            and not V2C_STOP_OPTION_RE.match(t_stem)
        )
        if dangling and (i + 1) < n:
            nxt = v2c_clean_text(items[i + 1].get("text", ""))
            nxt_stem = v2c_strip_q_prefix(nxt)
            if (
                nxt_stem
                and nxt_stem[:1].islower()
                and not V2C_LETTERED_OPT_PREFIX_RE.match(nxt_stem)
                and not _v2c_looks_like_matching_stem(nxt_stem)
                and not V2C_NOISE_RE.match(nxt_stem)
                and not V2C_STOP_OPTION_RE.match(nxt_stem)
            ):
                combined = v2c_clean_text(f"{t_stem} {nxt_stem}")
                if "?" not in combined:
                    combined = combined.rstrip(".") + "?"
                out.append({"text": combined, "is_red": False})
                i += 2
                continue
        out.append(it)
        i += 1
    return out


def v2c_parse_mcq_questions(items: list[dict]) -> list[dict]:
    questions_list: list[dict] = []
    current_q: str | None = None
    current_opts: list[dict] = []
    saw_multi_hint = False
    current_start_idx: int | None = None
    pending_multi_hint = False

    instruction_block_re = re.compile(r"^(instructions|for\s+learners|for\s+students|for\s+assessors)\b", re.IGNORECASE)
    meta_line_re = re.compile(r"^(More than one answer may apply|Select all that apply|Choose all that apply)\b", re.IGNORECASE)
    meta_any_re = re.compile(r"\b(More than one answer may apply|Select all that apply|Choose all that apply)\b", re.IGNORECASE)
    colon_stem_re = re.compile(r":\s*(?:\((?:select|choose)\b.*\))?\s*$", re.IGNORECASE)
    question_start_re = re.compile(r"^(?:in\s+)?(which|what|why|how)\b", re.IGNORECASE)
    select_stem_re = re.compile(r"^(?:q\s*\d+\.?\s*)?(select|choose|pick)\s+the\s+(best|correct|most\s+appropriate)\b", re.IGNORECASE)
    read_stem_re = re.compile(r"^(?:q\s*\d+\.?\s*)?read\s+the\s+following\b", re.IGNORECASE)
    complete_stem_re = re.compile(r"^(?:q\s*\d+\.?\s*)?complete\s+the\b", re.IGNORECASE)
    select_hint_re = re.compile(r"\((select|choose)\b", re.IGNORECASE)
    fill_gap_block_re = re.compile(r"\bfill\s+the\s+(gap|blank)\b", re.IGNORECASE)
    contains_select_summary_re = re.compile(r"\b(select|choose|pick)\s+the\s+(best|correct|most\s+appropriate)\s+summary\b", re.IGNORECASE)
    best_match_re = re.compile(r"\b(best\s+match|does\s+the\s+following\s+description\s+best\s+match)\b", re.IGNORECASE)

    def is_strong_stem(txt: str) -> bool:
        t = (txt or "").strip()
        return bool(
            t.endswith("?")
            or V2C_QUESTION_CMD_INNER_RE.search(t)
            or select_hint_re.search(t)
            or meta_any_re.search(t)
            or colon_stem_re.search(t)
            or select_stem_re.match(t)
            or question_start_re.match(t)
        )

    def flush():
        nonlocal current_q, current_opts, saw_multi_hint, current_start_idx
        if not current_q:
            return
        opts = [o for o in current_opts if not V2C_NOISE_RE.match(o["text"]) and not V2C_OPTION_NOISE_RE.match(o["text"])]
        option_texts = [o["text"] for o in opts]
        correct = [i for i, o in enumerate(opts) if o["is_red"]]
        qtext = v2c_strip_q_prefix(current_q.strip())
        qlower = qtext.lower()
        multi = (
            saw_multi_hint
            or bool(re.search(r"\bselect\s+(two|three|four|five|\d+)", qlower))
            or ("apply" in qlower)
            or (len(correct) > 1)
        )
        if len(option_texts) < 2:
            if not is_strong_stem(qtext):
                current_q = None
                current_opts = []
                saw_multi_hint = False
                current_start_idx = None
                return
            option_texts = [
                "⚠ Option text not extracted (likely image/shape). Please replace this option.",
                "⚠ Option text not extracted (likely image/shape). Please replace this option.",
            ]
            correct = []
        questions_list.append({
            "question": qtext,
            "options": option_texts,
            "correct": correct,
            "multi": multi,
            "kind": "mcq",
            "_order": current_start_idx if current_start_idx is not None else 10**9,
            "qnum": None,
        })
        current_q = None
        current_opts = []
        saw_multi_hint = False
        current_start_idx = None

    def parse_fill_gap_line(line: str):
        if line.count("/") < 2:
            return None
        parts = [p.strip() for p in re.split(r"\s*/\s*", line) if p.strip()]
        if len(parts) < 3:
            return None
        opt0 = parts[0].split()[-1]
        opt_last = parts[-1].split()[0]
        if not opt0 or not opt_last:
            return None
        prefix = parts[0][: -len(opt0)].rstrip()
        suffix = parts[-1][len(opt_last):].lstrip()
        options = [opt0] + parts[1:-1] + [opt_last]
        options = [v2c_clean_text(o) for o in options if v2c_clean_text(o)]
        qtext = v2c_clean_text(f"{prefix} ____ {suffix}".strip())
        if len(qtext) < 10 or len(options) < 3:
            return None
        return qtext, options

    def has_plausible_options(start_idx: int) -> bool:
        n = len(items)
        count = 0
        for j in range(start_idx, min(n, start_idx + 25)):
            raw = v2c_clean_text(items[j].get("text", ""))
            if not raw or V2C_NOISE_RE.match(raw):
                continue
            t = v2c_strip_q_prefix(raw)
            if V2C_ESSAY_GUIDE_RE.match(t) or V2C_RUBRIC_START_RE.match(t):
                return False
            if _v2c_looks_like_matching_stem(t):
                return False
            if select_stem_re.match(t) or V2C_QUESTION_CMD_INNER_RE.search(t) or best_match_re.search(t) or contains_select_summary_re.search(t):
                break
            if select_hint_re.search(t):
                break
            if t.endswith("?") and len(t) >= 10:
                break
            if len(t) <= 200 and not t.endswith("?"):
                count += 1
                if count >= 2:
                    return True
        return False

    for idx, it in enumerate(items):
        line = v2c_clean_text(it.get("text", ""))
        if not line or V2C_NOISE_RE.match(line) or V2C_OPTION_NOISE_RE.match(line):
            continue
        if instruction_block_re.match(v2c_strip_q_prefix(line)):
            flush()
            current_q = None
            current_opts = []
            saw_multi_hint = False
            current_start_idx = None
            pending_multi_hint = False
            continue
        if V2C_ESSAY_GUIDE_RE.match(line):
            current_q = None
            current_opts = []
            continue
        if current_q and V2C_STOP_OPTION_RE.match(line):
            flush()
            current_q = None
            current_opts = []
            continue

        t_stem = v2c_strip_q_prefix(line)

        if current_q is None and meta_line_re.match(t_stem):
            pending_multi_hint = True
            continue
        if _v2c_looks_like_matching_stem(t_stem):
            flush()
            current_q = None
            current_opts = []
            saw_multi_hint = False
            current_start_idx = None
            continue

        if (
            (current_q is None or len(current_opts) >= 2)
            and re.search(r":\s*(?:\((?:select|choose)\b.*\))?\s*$", t_stem, re.IGNORECASE)
            and len(t_stem) >= 12
            and not _v2c_looks_like_matching_stem(t_stem)
            and not V2C_COMMAND_QUESTION_RE.match(t_stem)
            and not V2C_STOP_OPTION_RE.match(line)
            and has_plausible_options(idx + 1)
        ):
            flush()
            current_q = t_stem
            current_opts = []
            saw_multi_hint = bool(meta_any_re.search(t_stem) or pending_multi_hint)
            current_start_idx = idx
            pending_multi_hint = False
            continue

        if (
            (current_q is None or len(current_opts) >= 2)
            and meta_any_re.search(t_stem)
            and not meta_line_re.match(t_stem)
            and len(t_stem) >= 12
            and not _v2c_looks_like_matching_stem(t_stem)
            and not V2C_COMMAND_QUESTION_RE.match(t_stem)
            and not V2C_STOP_OPTION_RE.match(line)
        ):
            flush()
            current_q = t_stem
            current_opts = []
            saw_multi_hint = True
            current_start_idx = idx
            pending_multi_hint = False
            continue

        if fill_gap_block_re.search(t_stem):
            flush()
            current_q = None
            current_opts = []
            saw_multi_hint = False
            current_start_idx = None
            continue

        if (current_q is None) and "/" in t_stem and has_plausible_options(idx + 1):
            parsed = parse_fill_gap_line(t_stem)
            if parsed:
                qtext, opts = parsed
                questions_list.append({
                    "question": qtext,
                    "options": opts,
                    "correct": [],
                    "multi": False,
                    "kind": "mcq",
                    "_order": idx,
                    "qnum": None,
                })
                continue

        if select_hint_re.search(t_stem) and (current_q is None or len(current_opts) >= 2) and has_plausible_options(idx + 1):
            flush()
            current_q = t_stem
            current_opts = []
            saw_multi_hint = pending_multi_hint
            current_start_idx = idx
            pending_multi_hint = False
            continue

        if (
            (read_stem_re.match(t_stem) or complete_stem_re.match(t_stem))
            and ("select" in t_stem.lower() or "most appropriate" in t_stem.lower() or "complete" in t_stem.lower())
            and not _v2c_looks_like_matching_stem(t_stem)
            and (current_q is None or len(current_opts) >= 2)
            and has_plausible_options(idx + 1)
        ):
            flush()
            current_q = t_stem
            current_opts = []
            saw_multi_hint = pending_multi_hint
            current_start_idx = idx
            pending_multi_hint = False
            continue

        if select_stem_re.match(line) and not _v2c_looks_like_matching_stem(line) and not meta_line_re.match(line):
            flush()
            current_q = t_stem
            current_opts = []
            saw_multi_hint = pending_multi_hint
            current_start_idx = idx
            pending_multi_hint = False
            continue

        m = V2C_QUESTION_CMD_INNER_RE.search(line)
        if m:
            flush()
            start = m.start()
            stem = line[:start].strip()
            cmd_plus = line[start:].strip()
            q_line = f"{stem} {cmd_plus}".strip() if stem else cmd_plus
            current_q = v2c_strip_q_prefix(q_line)
            current_opts = []
            current_start_idx = idx
            saw_multi_hint = pending_multi_hint
            pending_multi_hint = False
            continue

        if current_q and not current_opts:
            if (
                not current_q.strip().endswith("?")
                and line[:1].islower()
                and not V2C_QUESTION_CMD_INNER_RE.search(line)
                and not _v2c_looks_like_matching_stem(line)
                and not V2C_COMMAND_QUESTION_RE.match(v2c_strip_q_prefix(line))
                and not V2C_STOP_OPTION_RE.match(line)
                and not meta_line_re.match(line)
            ):
                current_q = (current_q + " " + line).strip()
                continue

        if (
            question_start_re.match(t_stem)
            and len(t_stem) >= 12
            and not _v2c_looks_like_matching_stem(t_stem)
            and not V2C_COMMAND_QUESTION_RE.match(t_stem)
            and not V2C_STOP_OPTION_RE.match(line)
            and not meta_line_re.match(line)
            and (
                "?" in t_stem
                or best_match_re.search(t_stem)
                or contains_select_summary_re.search(t_stem)
                or re.search(r"\((select|choose)\b", t_stem, re.IGNORECASE)
            )
            and (current_q is None or len(current_opts) >= 2)
            and has_plausible_options(idx + 1)
        ):
            flush()
            current_q = t_stem
            current_opts = []
            saw_multi_hint = pending_multi_hint
            current_start_idx = idx
            pending_multi_hint = False
            continue

        if (contains_select_summary_re.search(t_stem) or best_match_re.search(t_stem)) and (current_q is None or len(current_opts) >= 2) and has_plausible_options(idx + 1):
            flush()
            current_q = t_stem
            current_opts = []
            saw_multi_hint = pending_multi_hint
            current_start_idx = idx
            pending_multi_hint = False
            continue

        if current_q and meta_line_re.match(line):
            saw_multi_hint = True
            continue

        t = v2c_strip_q_prefix(line)
        if t.endswith("?") and len(t) >= 10 and not V2C_COMMAND_QUESTION_RE.match(t) and not _v2c_looks_like_matching_stem(t) and has_plausible_options(idx + 1):
            flush()
            current_q = t
            current_opts = []
            current_start_idx = idx
            continue

        if current_q:
            current_opts.append({"text": line, "is_red": bool(it.get("is_red", False))})

    flush()
    return [q for q in questions_list if len(q.get("options") or []) >= 2 and len(q.get("question") or "") >= 10]


def v2c_parse_essay_questions(items: list[dict]) -> list[dict]:
    questions: list[dict] = []
    n = len(items)
    i = 0
    while i < n:
        raw = v2c_clean_text(items[i].get("text", ""))
        if not raw or V2C_NOISE_RE.match(raw):
            i += 1
            continue
        line = v2c_strip_q_prefix(raw)
        if V2C_COMMAND_QUESTION_RE.match(line):
            j = i + 1
            next_line = ""
            while j < n:
                nxt = v2c_clean_text(items[j].get("text", ""))
                if nxt and not V2C_NOISE_RE.match(nxt):
                    next_line = nxt
                    break
                j += 1
            if V2C_RUBRIC_START_RE.match(next_line):
                questions.append({
                    "question": line,
                    "options": [],
                    "correct": [],
                    "multi": False,
                    "kind": "essay",
                    "_order": i,
                    "qnum": None,
                })
                i = j + 1
                continue
        i += 1
    return [q for q in questions if len((q.get("question") or "").strip()) >= 10]


# ===========================================================================
# V3 Essay parser (rule-based fallback)
# ===========================================================================
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