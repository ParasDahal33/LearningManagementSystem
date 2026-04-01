"""
services/openai_service.py
OpenAI API integration: config dataclass, JSON-schema structured calls,
and AI-based question segmentation for v2 and v3 parsers.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass

import requests

from core.utils import (
    normalize_key,
    v2_clean_text,
    v2_normalize_key,
    v2_strip_q_prefix,
    v2_strip_answer_guide,
    v2_trim_after_question_mark,
    V2_ANSWER_GUIDE_START_RE,
    v3_clean_text,
    v3_normalize_key,
    v3_strip_q_prefix,
    v3_strip_answer_guide,
    v3_trim_after_question_mark,
    v3_trim_after_sentence_if_long,
    V3_ANSWER_GUIDE_START_RE,
    V3_ANSWER_GUIDE_ANY_RE,
)
from parsers.mcq_parsers import (
    V3_IGNORE_LINE_RE,
    V3_IGNORE_TABLE_RE,
)


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
@dataclass
class OpenAIConfig:
    api_key: str
    model: str
    base_url: str = "https://api.openai.com"
    timeout_s: int = 120


# ---------------------------------------------------------------------------
# Low-level API call
# ---------------------------------------------------------------------------
def openai_responses_json_schema(
    prompt: str,
    schema_name: str,
    schema: dict,
    cfg: OpenAIConfig,
) -> tuple[dict | None, str | None]:
    """
    POST to the OpenAI /v1/responses endpoint with a JSON-schema output format.
    Returns (parsed_dict, None) on success or (None, error_message) on failure.
    """
    url = cfg.base_url.rstrip("/") + "/v1/responses"
    headers = {"Authorization": f"Bearer {cfg.api_key}", "Content-Type": "application/json"}
    body = {
        "model": cfg.model,
        "input": prompt,
        "text": {
            "format": {
                "type": "json_schema",
                "name": schema_name,
                "schema": schema,
                "strict": True,
            }
        },
    }
    try:
        r = requests.post(url, headers=headers, json=body, timeout=cfg.timeout_s)
    except Exception as e:
        return None, f"OpenAI request failed: {e}"
    if r.status_code >= 400:
        return None, f"OpenAI error {r.status_code}: {r.text}"
    try:
        data = r.json()
    except Exception as e:
        return None, f"OpenAI JSON parse failed: {e}"
    try:
        out = data["output"][0]["content"][0]
        if out.get("type") == "output_text" and out.get("text"):
            return json.loads(out["text"]), None
        if out.get("type") == "output_json" and out.get("json"):
            return out["json"], None
        if "text" in out:
            return json.loads(out["text"]), None
    except Exception as e:
        return None, f"OpenAI response parse failed: {e}"
    return None, "OpenAI returned an unexpected response shape."


# ---------------------------------------------------------------------------
# Shared schema
# ---------------------------------------------------------------------------
_SEGMENT_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "questions": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "kind": {"type": "string", "enum": ["mcq", "essay"]},
                    "stem": {"type": "array", "items": {"type": "integer"}},
                    "options": {"type": "array", "items": {"type": "array", "items": {"type": "integer"}}},
                },
                "required": ["kind", "stem", "options"],
            },
        }
    },
    "required": ["questions"],
}

''' _BASE_PROMPT = (
    "You are segmenting a DOCX extraction into Canvas quiz questions.\n"
    "Return STRICT JSON only (per schema).\n"
    "\n"
    "Hard rules:\n"
    "- You MUST NOT invent any text.\n"
    "- You may ONLY reference item indices (I<n>) from the provided list.\n"
    "- Keep original order (earlier indices first).\n"
    "- Do NOT create questions from instructions/policy/rubric.\n"
    "- Do NOT include assessor guide content like 'Answer may/must/needs address' or sample answers in stems/options.\n"
    "- For MCQ: include ALL options (both R0 and R1). Mark correctness by whether an option contains any R1 items.\n"
    "- For essay questions: options must be [].\n"
    "\n"
    "MCQ rules:\n"
    "- Options are typically lettered (a), (b), etc or separate lines under a prompt.\n"
    "- If you cannot find at least 2 options, do NOT output an MCQ.\n"
    "\n"
    "Essay rules:\n"
    "- Use the question prompt only.\n"
    "\n"
    "Items (format: I<index>|R0/R1|text):\n"
)'''

_BASE_PROMPT = (
    "You are an expert instructional data parser segmenting a DOCX extraction into Canvas LMS quiz questions.\n"
    "Return STRICT, valid JSON only, following the exact schema provided. Do not include markdown formatting like ```json.\n"
    "\n"
    "HARD RULES (Failure to follow these will break the system):\n"
    "- 1. NO HALLUCINATIONS: You MUST NOT invent, rephrase, or summarize any text. Use the exact text provided.\n"
    "- 2. INDEX TRACKING: You may ONLY reference item indices (e.g., I5) from the provided list. Keep the original chronological order.\n"
    "- 3. EXCLUSIONS: Ignore general document instructions, policies, and table headers. HOWEVER, you MUST retain Assessor notes (e.g., 'Answer may address...') specifically attached to short answer/essay questions.\n"
    "- 4. EXACT SCHEMA: Your output must map directly to the provided JSON structure.\n"
    "\n"
    "QUESTION TYPE RULES:\n"
    "- MULTIPLE CHOICE (single correct): If only one option has an R1 tag, set 'question_type' to 'multiple_choice_question'.\n"
    "- MULTIPLE ANSWERS (multi-select): If the question asks to 'Select two/three/four' OR if multiple options have an R1 tag, set 'question_type' to 'multiple_answers_question'.\n"
    "- ESSAY/SHORT ANSWER: If the question is open-ended and followed by assessor grading notes rather than selectable options, set 'question_type' to 'essay_question'.\n"
    "\n"
    "MCQ OPTION & ANSWER RULES:\n"
    "- Extract ALL available options associated with the question stem.\n"
    "- Inside the 'answers' array, map correctness strictly using the R0/R1 tags (R1 = 100 weight, R0 = 0 weight).\n"
    "- Ensure 'neutral_comments' is left as an empty string (\"\").\n"
    "\n"
    "ESSAY/SHORT ANSWER RULES:\n"
    "- The 'answers' array MUST be completely empty: [].\n"
    "- Extract the assessor notes (e.g., 'Answer may address...') and all the associated bullet points (R1 tags) into a single formatted paragraph string. Map this string to the 'neutral_comments' field.\n"
    "\n"
    "Items to parse (format: I<index>|R0/R1|text):\n"
)


_MAX_BLOCK_ITEMS = 170
_OVERLAP = 50


def _build_blocks(n: int) -> list[tuple[int, int]]:
    blocks: list[tuple[int, int]] = []
    start = 0
    while start < n:
        end = min(n, start + _MAX_BLOCK_ITEMS)
        blocks.append((start, end))
        if end >= n:
            break
        start = max(0, end - _OVERLAP)
    return blocks


# ===========================================================================
# V2 AI segmentation
# ===========================================================================
def v2_ai_segment_items_openai(
    items: list[dict], cfg: OpenAIConfig
) -> tuple[list[dict], list[str]]:
    log: list[str] = []
    if not items:
        return [], log

    def to_line(i: int) -> str:
        t = v2_clean_text(items[i].get("text", ""))
        red = "R1" if items[i].get("is_red") else "R0"
        return f"I{i}|{red}|{t}"

    all_qs: list[dict] = []

    for a, b in _build_blocks(len(items)):
        ctx = [to_line(i) for i in range(a, b) if v2_clean_text(items[i].get("text", ""))]
        if len(ctx) < 6:
            continue
        log.append(f"AI block: {a}-{b} lines={len(ctx)}")
        prompt = _BASE_PROMPT.replace(
            "- For MCQ: include ALL options (both R0 and R1). Mark correctness by whether an option contains any R1 items.\n"
            "- For essay questions: options must be [].\n",
            "- Correct MCQ options are those with R1 (red). For essay questions: options must be [].\n",
        ) + "\n".join(ctx)

        data, err = openai_responses_json_schema(prompt, "segment_questions", _SEGMENT_SCHEMA, cfg)
        if err:
            log.append(f"  block failed: {err}")
            continue
        qs = data.get("questions") if isinstance(data, dict) else None
        if not isinstance(qs, list):
            log.append("  block skipped: missing questions[]")
            continue

        for q in qs:
            if not isinstance(q, dict):
                continue
            kind = (q.get("kind") or "").strip().lower()
            stem_ids = q.get("stem") if isinstance(q.get("stem"), list) else []
            if kind not in ("mcq", "essay") or not stem_ids:
                continue
            if not all(isinstance(x, int) and 0 <= x < len(items) for x in stem_ids):
                continue

            stem_text = v2_clean_text(" ".join(v2_clean_text(items[x].get("text", "")) for x in stem_ids))
            stem_text = v2_strip_q_prefix(v2_strip_answer_guide(stem_text))
            stem_text = v2_trim_after_question_mark(stem_text)
            if not stem_text or len(stem_text) < 10:
                continue
            if stem_text.lower().startswith(("answer may address", "answer must address", "answer needs to address")):
                continue

            if kind == "essay":
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids), "qnum": None,
                })
                continue

            opt_groups = q.get("options") if isinstance(q.get("options"), list) else []
            option_texts: list[str] = []
            correct: list[int] = []
            for group in opt_groups:
                if not isinstance(group, list) or not group:
                    continue
                if not all(isinstance(x, int) and 0 <= x < len(items) for x in group):
                    continue
                t = v2_clean_text(" ".join(v2_clean_text(items[x].get("text", "")) for x in group))
                if not t or V2_ANSWER_GUIDE_START_RE.match(t):
                    continue
                option_texts.append(t)
                if any(bool(items[x].get("is_red")) for x in group):
                    correct.append(len(option_texts) - 1)

            seen: set[str] = set()
            out_opts: list[str] = []
            out_corr: list[int] = []
            for i_opt, opt in enumerate(option_texts):
                k = v2_normalize_key(opt)
                if k in seen:
                    continue
                seen.add(k)
                if i_opt in correct:
                    out_corr.append(len(out_opts))
                out_opts.append(opt)
            if len(out_opts) < 2:
                continue

            all_qs.append({
                "question": stem_text,
                "options": out_opts,
                "correct": out_corr,
                "multi": ("apply" in stem_text.lower()) or (len(out_corr) > 1),
                "kind": "mcq",
                "_order": min(stem_ids),
                "qnum": None,
            })

    all_qs.sort(key=lambda q: int(q.get("_order", 10**9)))
    return all_qs, log


# ===========================================================================
# V3 AI segmentation
# ===========================================================================
def v3_ai_segment_items_openai(
    items: list[dict], cfg: OpenAIConfig
) -> tuple[list[dict], list[str]]:
    log: list[str] = []
    if not items:
        return [], log

    def to_line(i: int) -> str:
        t = v3_clean_text(items[i].get("text", ""))
        red = "R1" if items[i].get("is_red") else "R0"
        return f"I{i}|{red}|{t}"

    def _v3_looks_like_question_start(text: str) -> bool:
        try:
            from parsers.mcq_parsers import _v3_looks_like_question_start as _inner
            return _inner(text)
        except ImportError:
            return True

    def should_demote_mcq_to_essay(stem_text: str, options: list[str], correct: list[int]) -> bool:
        s = v3_normalize_key(stem_text)
        if not s or len(options) <= 1:
            return True
        mcq_cue = bool(re.search(
            r"\b(which of the following|which strategy or technique|stand for|select|choose|pick|"
            r"more than one answer|select all that apply|choose all that apply)\b", s
        ))
        if not correct and not mcq_cue:
            return True
        if len(options) == 2:
            a, b = v3_normalize_key(options[0]), v3_normalize_key(options[1])
            if a and b and (a in b or b in a):
                return True
        if s.startswith(("what is the name", "what was the name", "what is meant", "what is the origin")):
            if len(options) <= 3 and len(correct) <= 1:
                return True
        return any(
            v3_normalize_key(opt) and len(v3_normalize_key(opt)) > 25
            and (v3_normalize_key(opt) in s or s in v3_normalize_key(opt))
            for opt in options
        )

    def _looks_like_continuation(text: str) -> bool:
        t2 = (text or "").strip()
        return not t2 or t2[:1].islower() or t2.startswith(("•", "-", "–", "—", ",", ";", ":", ")", "]"))

    def _split_group_into_segments(idxs: list[int]) -> list[list[int]]:
        if len(idxs) <= 1:
            return [idxs]
        segments: list[list[int]] = []
        cur: list[int] = []
        for ix in idxs:
            tx = v3_clean_text(items[ix].get("text", ""))
            if not cur:
                cur = [ix]
                continue
            if _looks_like_continuation(tx):
                cur.append(ix)
            else:
                segments.append(cur)
                cur = [ix]
        if cur:
            segments.append(cur)
        return [idxs] if len(segments) > 6 else segments

    all_qs: list[dict] = []

    for a, b in _build_blocks(len(items)):
        ctx = [to_line(i) for i in range(a, b) if v3_clean_text(items[i].get("text", ""))]
        if len(ctx) < 6:
            continue
        log.append(f"AI block: {a}-{b} lines={len(ctx)}")
        prompt = _BASE_PROMPT + "\n".join(ctx)

        data, err = openai_responses_json_schema(prompt, "segment_questions", _SEGMENT_SCHEMA, cfg)
        if err:
            log.append(f"  block failed: {err}")
            continue
        qs = data.get("questions") if isinstance(data, dict) else None
        if not isinstance(qs, list):
            log.append("  block skipped: missing questions[]")
            continue

        for q in qs:
            if not isinstance(q, dict):
                continue
            kind = (q.get("kind") or "").strip().lower()
            stem_ids = q.get("stem") if isinstance(q.get("stem"), list) else []
            if kind not in ("mcq", "essay") or not stem_ids:
                continue
            if not all(isinstance(x, int) and 0 <= x < len(items) for x in stem_ids):
                continue

            stem_text = v3_clean_text(" ".join(v3_clean_text(items[x].get("text", "")) for x in stem_ids))
            stem_text = v3_strip_q_prefix(v3_strip_answer_guide(stem_text))
            stem_text = v3_trim_after_question_mark(stem_text)
            stem_text = v3_trim_after_sentence_if_long(stem_text)
            if not stem_text or len(stem_text) < 10:
                continue
            if stem_text.lower().startswith(("answer may address", "answer must address", "answer needs to address")):
                continue
            if not _v3_looks_like_question_start(stem_text):
                continue

            if kind == "essay":
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids), "qnum": None,
                })
                continue

            opt_groups = q.get("options") if isinstance(q.get("options"), list) else []
            option_texts: list[str] = []
            correct: list[int] = []
            for group in opt_groups:
                if not isinstance(group, list) or not group:
                    continue
                if not all(isinstance(x, int) and 0 <= x < len(items) for x in group):
                    continue
                for seg in _split_group_into_segments(group):
                    t = v3_clean_text(" ".join(v3_clean_text(items[x].get("text", "")) for x in seg))
                    if not t or V3_ANSWER_GUIDE_START_RE.match(t) or V3_IGNORE_TABLE_RE.match(t) or V3_IGNORE_LINE_RE.match(t):
                        continue
                    option_texts.append(t)
                    if any(bool(items[x].get("is_red")) for x in seg):
                        correct.append(len(option_texts) - 1)

            seen: set[str] = set()
            out_opts: list[str] = []
            out_corr: list[int] = []
            for i_opt, opt in enumerate(option_texts):
                k = v3_normalize_key(opt)
                if k in seen:
                    continue
                seen.add(k)
                if i_opt in correct:
                    out_corr.append(len(out_opts))
                out_opts.append(opt)
            if len(out_opts) < 2:
                continue

            if should_demote_mcq_to_essay(stem_text, out_opts, out_corr):
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids), "qnum": None,
                })
            else:
                all_qs.append({
                    "question": stem_text,
                    "options": out_opts,
                    "correct": out_corr,
                    "multi": ("apply" in stem_text.lower()) or (len(out_corr) > 1),
                    "kind": "mcq",
                    "_order": min(stem_ids),
                    "qnum": None,
                })

    deduped: list[dict] = []
    seen_q: set[str] = set()
    for q in sorted(all_qs, key=lambda q: int(q.get("_order", 10**9))):
        k = v3_normalize_key(q.get("question", ""))
        if not k or k in seen_q:
            continue
        seen_q.add(k)
        deduped.append(q)
    return deduped, log