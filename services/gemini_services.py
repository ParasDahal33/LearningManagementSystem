"""
services/gemini_services.py
Google Gemini API integration: config dataclass, structured JSON calls,
and AI-based question segmentation mirroring the openai_services logic.
"""

from __future__ import annotations

import json
import re
import requests
from dataclasses import dataclass

from core.utils import (
    v3_clean_text,
    v3_normalize_key,
    v3_strip_q_prefix,
    v3_strip_answer_guide,
    v3_trim_after_question_mark,
    v3_trim_after_sentence_if_long,
    V3_ANSWER_GUIDE_START_RE,
    strip_q_prefix,
)
from parsers.mcq_parsers import (
    V3_IGNORE_LINE_RE,
    V3_IGNORE_TABLE_RE,
)

# We try to import this but provide a fallback if the module structure is slightly different
try:
    from parsers.mcq_parsers import _v3_looks_like_question_start
except ImportError:
    def _v3_looks_like_question_start(text: str) -> bool: return True

@dataclass
class GeminiConfig:
    api_key: str
    model: str = "gemini-1.5-flash"
    base_url: str = "https://generativelanguage.googleapis.com"
    timeout_s: int = 120


def gemini_responses_json_schema(
    prompt: str,
    schema: dict,
    cfg: GeminiConfig,
) -> tuple[dict | None, str | None]:
    """
    Performs a POST request to the Google Gemini API using structured JSON output mode.
    Returns (parsed_dict, None) on success or (None, error_message) on failure.
    """
    url = f"{cfg.base_url.rstrip('/')}/v1beta/models/{cfg.model}:generateContent?key={cfg.api_key}"
    headers = {"Content-Type": "application/json"}
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "response_mime_type": "application/json",
            "response_schema": schema
        }
    }
    try:
        r = requests.post(url, headers=headers, json=body, timeout=cfg.timeout_s)
        if r.status_code >= 400:
            return None, f"Gemini API error {r.status_code}: {r.text}"
        
        data = r.json()
        if "candidates" not in data or not data["candidates"]:
            return None, "Gemini returned no candidates."
        
        cand = data["candidates"][0]
        if "content" not in cand or "parts" not in cand["content"]:
            return None, "Gemini response structure invalid."
            
        text = cand["content"]["parts"][0]["text"]
        return json.loads(text), None
    except Exception as e:
        return None, f"Gemini request failed: {e}"


_SEGMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "questions": {
            "type": "array",
            "items": {
                "type": "object",
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

_BASE_PROMPT = (
    "You are an expert at segmenting extracted DOCX text into Canvas quiz questions.\n"
    "Return STRICT JSON only following the provided schema.\n"
    "\n"
    "Rules:\n"
    "- Use ONLY the item indices (I<n>) provided.\n"
    "- Do NOT invent new text.\n"
    "- For MCQ: include all options. Options containing R1 items are correct.\n"
    "- For Essay: options must be [].\n"
    "\n"
    "Items (Format: I<index>|R0/R1|text):\n"
)

_MAX_BLOCK_ITEMS = 170
_OVERLAP = 50

def _build_blocks(n: int) -> list[tuple[int, int]]:
    blocks: list[tuple[int, int]] = []
    start = 0
    while start < n:
        end = min(n, start + _MAX_BLOCK_ITEMS)
        blocks.append((start, end))
        if end >= n: break
        start = max(0, end - _OVERLAP)
    return blocks

def _should_demote_mcq_to_essay(stem_text: str, options: list[str], correct: list[int]) -> bool:
    s = v3_normalize_key(stem_text)
    if not s or len(options) <= 1:
        return True
    mcq_cue = bool(re.search(
        r"\b(which of the following|which strategy|stand for|select|choose|pick|select all that apply)\b", s
    ))
    if not correct and not mcq_cue:
        return True
    if len(options) == 2:
        a, b = v3_normalize_key(options[0]), v3_normalize_key(options[1])
        if a and b and (a in b or b in a): return True
    return False

def v3_ai_segment_items_gemini(
    items: list[dict], cfg: GeminiConfig
) -> tuple[list[dict], list[str]]:
    log: list[str] = []
    if not items:
        return [], log

    def to_line(i: int) -> str:
        t = v3_clean_text(items[i].get("text", ""))
        red = "R1" if items[i].get("is_red") else "R0"
        return f"I{i}|{red}|{t}"

    all_qs: list[dict] = []
    blocks = _build_blocks(len(items))
    
    for a, b in blocks:
        ctx = [to_line(i) for i in range(a, b) if v3_clean_text(items[i].get("text", ""))]
        if len(ctx) < 5: continue
        
        log.append(f"AI block (Gemini): {a}-{b} lines={len(ctx)}")
        prompt = _BASE_PROMPT + "\n".join(ctx)

        data, err = gemini_responses_json_schema(prompt, _SEGMENT_SCHEMA, cfg)
        if err:
            log.append(f"  block failed: {err}")
            continue
            
        qs_list = data.get("questions", []) if isinstance(data, dict) else []
        for q_data in qs_list:
            kind = (q_data.get("kind") or "").strip().lower()
            stem_ids = q_data.get("stem", [])
            if not stem_ids or not all(isinstance(x, int) and 0 <= x < len(items) for x in stem_ids):
                continue

            # Reconstruct and clean stem
            stem_text = v3_clean_text(" ".join(items[x].get("text", "") for x in stem_ids))
            stem_text = v3_strip_q_prefix(v3_strip_answer_guide(stem_text))
            stem_text = v3_trim_after_question_mark(stem_text)
            stem_text = v3_trim_after_sentence_if_long(stem_text)
            
            if not stem_text or len(stem_text) < 8: continue
            if not _v3_looks_like_question_start(stem_text): continue

            if kind == "essay":
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids)
                })
                continue

            # Handle MCQ Options
            opt_groups = q_data.get("options", [])
            option_texts: list[str] = []
            correct_map: list[int] = []
            
            for group in opt_groups:
                if not group: continue
                t = v3_clean_text(" ".join(items[x].get("text", "") for x in group))
                if not t or V3_ANSWER_GUIDE_START_RE.match(t) or V3_IGNORE_TABLE_RE.match(t) or V3_IGNORE_LINE_RE.match(t):
                    continue
                option_texts.append(t)
                if any(items[x].get("is_red") for x in group):
                    correct_map.append(len(option_texts) - 1)

            # Deduplicate options
            seen: set[str] = set()
            final_opts: list[str] = []
            final_corr: list[int] = []
            for idx, opt in enumerate(option_texts):
                k = v3_normalize_key(opt)
                if k and k not in seen:
                    seen.add(k)
                    if idx in correct_map: final_corr.append(len(final_opts))
                    final_opts.append(opt)

            if len(final_opts) < 2 or _should_demote_mcq_to_essay(stem_text, final_opts, final_corr):
                all_qs.append({
                    "question": stem_text, "options": [], "correct": [], "multi": False,
                    "kind": "essay", "_order": min(stem_ids)
                })
            else:
                all_qs.append({
                    "question": stem_text, "options": final_opts, "correct": final_corr,
                    "multi": (len(final_corr) > 1 or "apply" in stem_text.lower()),
                    "kind": "mcq", "_order": min(stem_ids)
                })

    # Final Global Deduplication
    deduped: list[dict] = []
    seen_questions: set[str] = set()
    for q in sorted(all_qs, key=lambda x: x["_order"]):
        key = v3_normalize_key(q["question"])
        if key not in seen_questions:
            seen_questions.add(key)
            deduped.append(q)
            
    return deduped, log