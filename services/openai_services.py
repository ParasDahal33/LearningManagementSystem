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
                    "neutral_comments": {"type": "string"},
                },
                # OpenAI's schema validator requires that every key listed in 'properties' be
                # included in 'required'. We include 'neutral_comments' here; the model may
                # return an empty string when no assessor notes exist.
                "required": ["kind", "stem", "options", "neutral_comments"],
            }
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


'''
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
     "MATCHING QUESTION RULES:\n"
    "- Detect matching sections when the source includes wording such as 'matching activity', 'Match each...', 'Column A', 'Column B', 'Answer', 'Ref', or an 'ASSESSOR KEY' containing number-letter mappings.\n"
    "- Treat each numbered entry in Column A as the prompt side of the match.\n"
    "- Treat each lettered entry in Column B as the choice side of the match.\n"
    "- Preserve the original numbered order from Column A and the original lettered order from Column B.\n"
    "- Use the exact text from Column A items as prompts and the exact text from Column B items as matches.\n"
    "- Derive correctness ONLY from the assessor key when present (for example: '1-E, 2-F, 3-G'). Do not infer or correct mappings.\n"
    "- Ignore instructional text such as 'Write the LETTER from Column B...' and 'Each letter ... is used only once per section.'\n"
    "- Ignore table labels and headers such as '#', 'Column A', 'Answer', 'Ref', and 'Column B'.\n"
    "- For matching questions, do not convert the content into multiple choice or essay format.\n"
    "- Populate the matching question exactly according to the target schema's matching structure. If the schema uses prompt/match pairs, each numbered Column A item must map to its correct lettered Column B item based strictly on the assessor key.\n"
    "- If a section contains multiple numbered prompts under one shared Column B list, treat that entire section as ONE matching question, not separate standalone questions.\n"
    "- 'neutral_comments' must be an empty string (\"\") for matching questions unless the schema explicitly requires otherwise.\n"
    "\n"
    "REFERENCE PATTERN FOR MATCHING SECTIONS:\n"
    "- Matching sections may appear in forms like:\n"
    "  * 'Section X: ...'\n"
    "  * numbered items in Column A (1, 2, 3...)\n"
    "  * lettered options in Column B (A, B, C...)\n"
    "  * an assessor key such as 'ASSESSOR KEY: 1-E, 2-F, 3-G...'\n"
    "- These indicate a matching_question and must be parsed as one grouped matching item per section.\n"
    "\n"
    "Items to parse (format: I<index>|R0/R1|text):\n"
)
'''


_BASE_PROMPT = (
    "You are an expert assessment-question parser for Canvas LMS.\n"
    "You receive raw text extracted from DOCX assessment documents, including possible tables, answer blanks, assessor-only keys, rubrics, and formatting artifacts.\n"
    "Return STRICT, valid JSON only, matching the exact schema provided. Do not include markdown formatting such as ```json.\n"
    "\n"
    "CORE RULES:\n"
    "- NO HALLUCINATIONS: Do not invent, rephrase, summarize, or correct source text. Use the exact text provided.\n"
    "- INDEX TRACKING: Only reference item indices from the provided list, such as I5. Preserve the original chronological order.\n"
    "- LEARNER-FACING CONTENT ONLY: Parse learner-facing questions. Exclude cover pages, global instructions, policies, table headers, assessor signatures, dates, feedback sections, result fields, and other non-question content.\n"
    "- ASSESSOR CONTENT: Remove assessor-only content from question_text, but retain assessor notes, sample answers, rubrics, and acceptable-answer guidance in the appropriate schema fields.\n"
    "- EXACT SCHEMA: Output must map directly to the provided JSON schema. Use only fields and question_type values supported by that schema.\n"
    "\n"
    "QUESTION TEXT RULES:\n"
    "- Preserve the learner-facing wording exactly as question_text.\n"
    "- Remove answer blanks or labels such as 'Answer:', blank lines, underscores, and assessor-only answer keys from question_text.\n"
    "- Ignore table labels and headers such as '#', 'Column A', 'Column B', 'Answer', and 'Ref'.\n"
    "\n"
    "QUESTION TYPE RULES:\n"
    "- MULTIPLE CHOICE, SINGLE CORRECT: If the question has selectable options and only one option is explicitly correct, use the schema's single-answer MCQ type.\n"
    "- MULTIPLE ANSWERS / MULTI-SELECT: If the question asks to select multiple options, such as 'Select two/three/four' or 'select all that apply', or if multiple options are explicitly correct, use the schema's multi-answer MCQ type and set multiple_select=true if that field exists.\n"
    "- SHORT ANSWER / ESSAY: If the question is open-ended and followed by assessor guidance, sample answers, acceptable answers, or marking notes rather than selectable options, use the schema's short-answer or essay type.\n"
    "- MATCHING: If the source includes wording such as 'matching activity', 'Match each...', 'Column A', 'Column B', 'Answer', 'Ref', or an assessor key with number-letter mappings, use the schema's matching type.\n"
    "\n"
    "MCQ OPTION AND ANSWER RULES:\n"
    "- Extract all available options associated with the question stem.\n"
    "- Preserve option text exactly.\n"
    "- Normalise option keys to A, B, C... when labels exist. If labels are missing, assign A, B, C... in displayed order.\n"
    "- Determine correctness only from explicit source indicators, including R0/R1 tags or assessor keys such as 'ASSESSOR KEY: C' or 'Answer: C'.\n"
    "- Map R1 as correct and R0 as incorrect. If the schema uses weights, R1 = 100 and R0 = 0.\n"
    "- Do not guess correctness. If the source key is missing or ambiguous, leave correctness unset or false according to the schema, add a warning if supported, and lower confidence if supported.\n"
    "- Set neutral_comments to an empty string if required by the schema.\n"
    "\n"
    "SHORT ANSWER / ESSAY RULES:\n"
    "- The answers array must be empty if the schema uses an answers field for essay questions.\n"
    "- Capture assessor guidance such as 'Answer may address...', sample answers, rubrics, and associated bullet points exactly.\n"
    "- Put broad assessor guidance into neutral_comments, sample_answer, or marking_rubric according to the schema.\n"
    "- Put concise exact acceptable answers into acceptable_answers only when the source explicitly provides exact acceptable answer strings.\n"
    "- Do not split long rubrics or broad guidance into acceptable_answers.\n"
    "\n"
    "MATCHING QUESTION RULES:\n"
    "- Treat each matching section as one grouped matching question, not separate standalone questions.\n"
    "- Treat each numbered item in Column A as the prompt/left side.\n"
    "- Treat each lettered item in Column B as the choice/right side.\n"
    "- Preserve the original numbered order from Column A and the original lettered order from Column B.\n"
    "- Use exact Column A text as prompts and exact Column B text as matches, excluding Column B reference letters from the right-side text.\n"
    "- Derive correctness only from the assessor key when present, such as '1-E, 2-F, 3-G'. Do not infer, repair, or correct mappings.\n"
    "- If the assessor key appears inconsistent with visible text, follow the explicit key, add a warning if supported by the schema, and lower confidence if supported.\n"
    "- Ignore instructional text such as 'Write the LETTER from Column B...' and 'Each letter is used only once per section.'\n"
    "- Populate matching content according to the schema's matching structure, such as prompt/match pairs or left/right pairs.\n"
    "- Set neutral_comments to an empty string if required by the schema.\n"
    "\n"
    "POINTS AND CONFIDENCE RULES:\n"
    "- Use explicit marks or points if present. Otherwise default to 1 if the schema requires points.\n"
    "- Use confidence 0.95 or higher for clear formats with explicit keys.\n"
    "- Use confidence 0.70 to 0.90 when content is parseable but formatting is messy.\n"
    "- Use confidence below 0.70 when a key is missing, ambiguous, or inconsistent.\n"
    "\n"
    "REFERENCE PATTERN FOR MATCHING SECTIONS:\n"
    "- Matching sections may appear as section headings, numbered Column A items, lettered Column B options, and assessor keys such as 'ASSESSOR KEY: 1-E, 2-F, 3-G'.\n"
    "- These indicate one matching question per section.\n"
    "\n"
    "Items to parse, formatted as I<index>|R0/R1|text:\n"
)

_MAX_BLOCK_ITEMS = 170
_OVERLAP = 80


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
                # include any assessor/neutral comments produced by the AI in the parsed question
                nc = q.get("neutral_comments") if isinstance(q.get("neutral_comments"), str) else ""
                nc = v3_clean_text(nc) if nc else ""
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids), "qnum": None,
                    "assessor_key": nc or None,
                    "neutral_comments": nc or None,
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

    def _collect_assessor_notes_for_stem(stem_ids: list[int]) -> str:
        """Collect nearby R1/answer-guide lines after the stem to form assessor notes."""
        notes: list[str] = []
        if not stem_ids:
            return ""
        start = max(stem_ids) + 1
        end = min(len(items), start + 10)
        for i in range(start, end):
            it = items[i]
            t = v3_clean_text(it.get("text", "") or "")
            if not t:
                continue
            if it.get("is_red") or V3_ANSWER_GUIDE_ANY_RE.search(t) or V3_ANSWER_GUIDE_START_RE.match(t):
                notes.append(t)
                continue
            if re.search(r"\banswer (may|must|needs) address\b", t, flags=re.IGNORECASE):
                notes.append(t)
                continue
            if t.startswith(("•", "-", "–", "—")):
                notes.append(t)
                continue
            if _v3_looks_like_question_start(t):
                break
        return "; ".join(notes).strip()

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
                nc = q.get("neutral_comments") if isinstance(q.get("neutral_comments"), str) else ""
                nc = v3_clean_text(nc) if nc else ""
                if not nc:
                    nc = _collect_assessor_notes_for_stem(stem_ids)
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids), "qnum": None,
                    "assessor_key": nc or None,
                    "neutral_comments": nc or None,
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
                nc = q.get("neutral_comments") if isinstance(q.get("neutral_comments"), str) else ""
                nc = v3_clean_text(nc) if nc else ""
                if not nc:
                    nc = _collect_assessor_notes_for_stem(stem_ids)
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids), "qnum": None,
                    "assessor_key": nc or None,
                    "neutral_comments": nc or None,
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


def v3_ai_extract_all_openai(items: list[dict], cfg: OpenAIConfig) -> tuple[list[dict], list[str]]:
    """Use OpenAI to extract MCQ, essay (short answer), and matching questions from items.

    The model must return a JSON array `questions` where each question has:
      - kind: 'mcq'|'essay'|'matching'
      - stem: [int,...]
      - options: [[int,...], ...]  (for mcq)
      - pairs: [{"left": int, "right": int}, ...] (for matching)
      - neutral_comments: string (may be empty)

    We post-process the indices into text and derive correctness from R1 (is_red) flags.
    """
    log: list[str] = []
    if not items:
        return [], log

    def to_line(i: int) -> str:
        t = v3_clean_text(items[i].get("text", ""))
        red = "R1" if items[i].get("is_red") else "R0"
        return f"I{i}|{red}|{t}"

    # JSON schema for the model response
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "questions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "kind": {"type": "string", "enum": ["mcq", "essay", "matching"]},
                        "stem": {"type": "array", "items": {"type": "integer"}},
                        "options": {"type": "array", "items": {"type": "array", "items": {"type": "integer"}}},
                        "pairs": {"type": "array", "items": {"type": "object", "properties": {"left": {"type": "integer"}, "right": {"type": "integer"}}, "required": ["left", "right"]}},
                        "neutral_comments": {"type": "string"},
                    },
                    "required": ["kind", "stem", "options", "pairs", "neutral_comments"],
                },
            }
        },
        "required": ["questions"],
    }

    # build blocks and call the API (reuse earlier block logic)
    all_qs: list[dict] = []
    for a, b in _build_blocks(len(items)):
        ctx = [to_line(i) for i in range(a, b) if v3_clean_text(items[i].get("text", ""))]
        if len(ctx) < 6:
            continue
        log.append(f"AI block: {a}-{b} lines={len(ctx)}")
        prompt = _BASE_PROMPT + "\n".join(ctx)
        data, err = openai_responses_json_schema(prompt, "segment_questions", schema, cfg)
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
            if kind not in ("mcq", "essay", "matching") or not stem_ids:
                continue
            if not all(isinstance(x, int) and 0 <= x < len(items) for x in stem_ids):
                continue

            # build stem text
            stem_text = v3_clean_text(" ".join(v3_clean_text(items[x].get("text", "")) for x in stem_ids))
            stem_text = v3_strip_q_prefix(v3_strip_answer_guide(stem_text))
            stem_text = v3_trim_after_question_mark(stem_text)
            stem_text = v3_trim_after_sentence_if_long(stem_text)
            if not stem_text or len(stem_text) < 8:
                continue

            nc = q.get("neutral_comments") if isinstance(q.get("neutral_comments"), str) else ""
            nc = v3_clean_text(nc) if nc else ""

            if kind == "mcq":
                opt_groups = q.get("options") if isinstance(q.get("options"), list) else []
                option_texts: list[str] = []
                correct: list[int] = []
                for group in opt_groups:
                    if not isinstance(group, list) or not group:
                        continue
                    t = v3_clean_text(" ".join(v3_clean_text(items[x].get("text", "")) for x in group))
                    if not t or V3_ANSWER_GUIDE_START_RE.match(t) or V3_IGNORE_TABLE_RE.match(t) or V3_IGNORE_LINE_RE.match(t):
                        continue
                    option_texts.append(t)
                    if any(bool(items[x].get("is_red")) for x in group):
                        correct.append(len(option_texts) - 1)

                if len(option_texts) < 2:
                    continue
                all_qs.append({
                    "question": stem_text,
                    "options": option_texts,
                    "correct": correct,
                    "multi": (len(correct) > 1),
                    "kind": "mcq",
                    "_order": min(stem_ids),
                    "qnum": None,
                    "assessor_key": nc or None,
                    "neutral_comments": nc or None,
                })

            elif kind == "essay":
                # fallback: if AI returned nothing, try to collect nearby R1 notes
                if not nc:
                    # look forward a few items for R1 / answer-guide lines
                    notes = []
                    start = max(stem_ids) + 1
                    for i in range(start, min(len(items), start + 8)):
                        t = v3_clean_text(items[i].get("text", "") or "")
                        if not t:
                            continue
                        if items[i].get("is_red") or V3_ANSWER_GUIDE_ANY_RE.search(t) or V3_ANSWER_GUIDE_START_RE.match(t):
                            notes.append(t)
                    nc = "; ".join(notes).strip()

                all_qs.append({
                    "question": stem_text,
                    "options": [],
                    "correct": [],
                    "multi": False,
                    "kind": "essay",
                    "_order": min(stem_ids),
                    "qnum": None,
                    "assessor_key": nc or None,
                    "neutral_comments": nc or None,
                })

            elif kind == "matching":
                pairs_in = q.get("pairs") if isinstance(q.get("pairs"), list) else []
                pairs: list[dict[str, str]] = []
                for p in pairs_in:
                    if not isinstance(p, dict):
                        continue
                    L = p.get("left")
                    R = p.get("right")
                    if not (isinstance(L, int) and isinstance(R, int)):
                        continue
                    if 0 <= L < len(items) and 0 <= R < len(items):
                        left_txt = v3_clean_text(items[L].get("text", ""))
                        right_txt = v3_clean_text(items[R].get("text", ""))
                        if left_txt and right_txt:
                            pairs.append({"left": left_txt, "right": right_txt})
                if not pairs:
                    continue
                all_qs.append({
                    "question": stem_text,
                    "pairs": pairs,
                    "kind": "matching",
                    "options": [],
                    "correct": [],
                    "multi": False,
                    "_order": min(stem_ids),
                    "qnum": None,
                    "assessor_key": nc or None,
                    "neutral_comments": nc or None,
                })

    # dedupe and return
    deduped: list[dict] = []
    seen_q: set[str] = set()
    for q in sorted(all_qs, key=lambda q: int(q.get("_order", 10**9))):
        k = v3_normalize_key(q.get("question", ""))
        if not k or k in seen_q:
            continue
        seen_q.add(k)
        deduped.append(q)
    return deduped, log