"""
ui/parse_section.py
Handles the DOCX upload and parsing UI logic.
"""
from __future__ import annotations

import contextlib
import io
import tempfile
import re
import json
import streamlit as st

from core.utils import dedupe_questions
from parsers.docx_extractor import extract_items_with_red_v1, build_description_v1, v2_extract_items_with_red, v3_extract_items_with_red
from services.openai_services import OpenAIConfig, v3_ai_segment_items_openai
from services.openai_services import rubrics_schema, rubrics_prompt, openai_responses_json_schema
from services.gemini_services import GeminiConfig, v3_ai_segment_items_gemini, gemini_responses_json_schema
from parsers.mcq_parsers import (
    parse_mcq_questions_v1,
    parse_essay_questions_v1,
    v2c_merge_dangling_question_lines,
    v2c_parse_mcq_questions,
    v2c_parse_essay_questions,
    v3_parse_essay_questions_rule_based,
    v3_filter_items_for_ai,
)
from parsers.matching_parser import (
    parse_matching_questions_doc_order_v1_exact,
    v3_parse_matching_questions_doc_order,
    v3_parse_table_defined_terms_as_essays,
    v3_parse_table_characteristics_as_essays,
    v3_collect_ignore_texts_from_forced_tables,
)
from core.utils import (
    v2_split_items_on_internal_qnums,
    v2c_clean_text,
    v2c_dedupe_questions,
    v2c_collapse_duplicate_mcq,
    v3_split_items_on_internal_qnums,
    v3_dedupe_questions,
)


def render_parse_section(parser_mode: str) -> None:
    st.subheader("1) Upload DOCX and Parse")
    uploaded = st.file_uploader("DOCX", type=["docx"], label_visibility="collapsed")

    c_parse1, c_parse2 = st.columns([1, 1])
    parse_btn = c_parse1.button("Parse", type="primary", use_container_width=True)
    clear_btn = c_parse2.button("Clear parsed results", use_container_width=True)
    log_box = st.empty()

    if clear_btn:
        st.session_state.questions = []
        st.session_state.parsed_ok = False
        st.session_state.description_html = ""
        st.session_state.docx_filename = None
        st.session_state.parse_run_id += 1
        st.rerun()

    if parse_btn:
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            if not uploaded:
                raise RuntimeError("Upload a DOCX to parse.")

            uploaded_name = (uploaded.name or "").strip() or "Quiz.docx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(uploaded.getvalue())
                docx_path = tmp.name

            st.session_state.docx_filename = uploaded_name

            # Description always built from v1 extraction (stable).
            desc_items = extract_items_with_red_v1(docx_path)
            st.session_state.description_html = build_description_v1(desc_items)

            qs: list[dict] = []
            ai_log: list[str] = []
            removed_dupes = 0

            # If user requested assessor guide structured import, call the Rubrics JSON-schema extractor
            if st.session_state.get("assessor_import"):
                # Build a doc text payload: use stable v1 extraction (paragraph order)
                text_items = extract_items_with_red_v1(docx_path)
                lines = [it.get("text", "") for it in text_items if it.get("text")]
                prompt = rubrics_prompt + "\n\nDOCUMENT:\n" + "\n".join(lines)

                ai_provider = st.session_state.get("ai_provider", "OpenAI")
                parsed = None
                if ai_provider == "OpenAI":
                    if not (st.session_state.openai_api_key or "").strip():
                        raise RuntimeError("Assessor import (rubrics) requires an OpenAI API key.")
                    cfg = OpenAIConfig(
                        api_key=st.session_state.openai_api_key.strip(),
                        model=(st.session_state.openai_model or "gpt-4o-mini").strip(),
                        base_url=(st.session_state.openai_base_url or "https://api.openai.com").strip(),
                    )
                    data, err = openai_responses_json_schema(prompt, "rubrics", rubrics_schema, cfg)
                    if err:
                        print("DEBUG: Rubrics extraction failed:", err)
                        st.error(f"Rubrics extraction failed: {err}")
                    else:
                        parsed = data
                else:
                    if not (st.session_state.gemini_api_key or "").strip():
                        raise RuntimeError("Assessor import (rubrics) requires a Gemini API key.")
                    cfg_g = GeminiConfig(
                        api_key=st.session_state.gemini_api_key.strip(),
                        model=(st.session_state.gemini_model or "gemini-1.5-flash").strip(),
                        base_url=(st.session_state.gemini_base_url or "https://generativelanguage.googleapis.com").strip(),
                    )
                    data, err = gemini_responses_json_schema(prompt, rubrics_schema, cfg_g)
                    if err:
                        print("DEBUG: Rubrics extraction (Gemini) failed:", err)
                        st.error(f"Rubrics extraction failed: {err}")
                    else:
                        parsed = data

                if parsed:
                    st.session_state.assessor_parse = parsed
                    st.session_state.questions = []
                    st.session_state.parsed_ok = True
                    st.session_state.parse_run_id += 1
                    # Log counts when possible
                    n_assign = len(parsed.get("assignments", [])) if isinstance(parsed.get("assignments"), list) else 0
                    print("DEBUG: Rubrics parsed: assignments=", n_assign)
            else:
                # ------------------------------------------------------------------ v1
                if parser_mode.startswith("v1"):
                    matching = parse_matching_questions_doc_order_v1_exact(docx_path)
                    mcq = parse_mcq_questions_v1(desc_items)
                    essay = parse_essay_questions_v1(desc_items)
                    qs = dedupe_questions(matching + mcq + essay)

                # ------------------------------------------------------------------ v2
                elif parser_mode.startswith("v2"):
                    items_v2 = v2c_merge_dangling_question_lines(desc_items)
                    items_v2 = v2_split_items_on_internal_qnums(items_v2)
                    matching = parse_matching_questions_doc_order_v1_exact(docx_path)
                    mcq = v2c_parse_mcq_questions(items_v2)
                    essay = v2c_parse_essay_questions(items_v2)
                    qs = v2c_dedupe_questions(matching + mcq + essay)
                    qs = v2c_collapse_duplicate_mcq(qs)

                # ------------------------------------------------------------------ v3
                else:
                    items = v3_extract_items_with_red(docx_path, include_tables=True)
                    items = v3_split_items_on_internal_qnums(items)

                    matching = v3_parse_matching_questions_doc_order(docx_path, items)
                    table_essays = v3_parse_table_defined_terms_as_essays(docx_path, items)
                    table_essays += v3_parse_table_characteristics_as_essays(docx_path, items)

                    ignore_terms: set[str] = set()
                    for q in table_essays:
                        qt = q.get("question", "")
                        m = re.match(r"^Define:\s*(.+?)\s*\.", qt, flags=re.IGNORECASE)
                        if m:
                            ignore_terms.add(m.group(1).strip())
                            continue
                        m = re.match(r"^Describe the essential characteristics of:\s*(.+?)\s*\.", qt, flags=re.IGNORECASE)
                        if m:
                            ignore_terms.add(m.group(1).strip())

                    ignore_texts = v3_collect_ignore_texts_from_forced_tables(docx_path)
                    ai_input = v3_filter_items_for_ai(items, ignore_terms=ignore_terms, ignore_texts=ignore_texts, mode="balanced")
                    
                    ai_provider = st.session_state.get("ai_provider", "OpenAI")
                    if ai_provider == "OpenAI":
                        if not (st.session_state.openai_api_key or "").strip():
                            raise RuntimeError("v3 (AI+fallback) requires an OpenAI API key.")
                        cfg = OpenAIConfig(
                            api_key=st.session_state.openai_api_key.strip(),
                            model=(st.session_state.openai_model or "gpt-4o-mini").strip(),
                            base_url=(st.session_state.openai_base_url or "https://api.openai.com").strip(),
                        )
                        ai_qs, ai_log = v3_ai_segment_items_openai(ai_input, cfg)
                    else:
                        if not (st.session_state.gemini_api_key or "").strip():
                            raise RuntimeError("v3 (AI+fallback) requires a Gemini API key.")
                        cfg_gemini = GeminiConfig(
                            api_key=st.session_state.gemini_api_key.strip(),
                            model=(st.session_state.gemini_model or "gemini-1.5-flash").strip(),
                            base_url=(st.session_state.gemini_base_url or "https://generativelanguage.googleapis.com").strip(),
                        )
                        ai_qs, ai_log = v3_ai_segment_items_gemini(ai_input, cfg_gemini)

                    rule_essays = v3_parse_essay_questions_rule_based(items)

                    qs = matching + table_essays + ai_qs + rule_essays
                    qs.sort(key=lambda q: int(q.get("_order", 10**9)))
                    qs, removed_dupes = v3_dedupe_questions(qs)

                st.session_state.questions = qs
                st.session_state.parsed_ok = True
                st.session_state.parse_run_id += 1
                st.session_state.details["quiz_title"] = ""
                st.session_state.details["quiz_instructions"] = ""

                print("DEBUG: parser:", parser_mode)
                print("DEBUG: items extracted (description v1):", len(desc_items))
                print("DEBUG: matching:", sum(1 for q in qs if q.get("kind") == "matching"))
                print("DEBUG: mcq:", sum(1 for q in qs if q.get("kind") == "mcq"))
                print("DEBUG: essay:", sum(1 for q in qs if q.get("kind") == "essay"))
                if parser_mode.startswith("v3"):
                    print("DEBUG: removed_dupes:", removed_dupes)
                print("Parsed questions:", len(qs))
                for ln in ai_log[:60]:
                    print(ln)

        log_box.code(buf.getvalue())
        st.success(f"✅ Parsed {len(st.session_state.questions)} questions.")

        # If we parsed an assessor guide (rubrics schema), show a preview panel
        if st.session_state.get("assessor_parse"):
            parsed = st.session_state.assessor_parse
            st.subheader("Assessor Guide  Parsed Preview")
            cols = st.columns([2, 1])
            with cols[0]:
                st.markdown("**Assignments / Stages**")
                assignments = parsed.get("assignments") if isinstance(parsed.get("assignments"), list) else []
                if not assignments:
                    st.info("No assignments found in parsed output.")
                for a in assignments:
                    title = a.get("title") or "(no title)"
                    with st.expander(title):
                        desc = a.get("description_html") or a.get("description") or ""
                        if desc:
                            st.markdown(desc, unsafe_allow_html=True)
                        sub_types = a.get("submission_types") or []
                        if sub_types:
                            st.caption("Submission types: " + ", ".join(sub_types))
                        rubric = a.get("rubric") or []
                        if rubric:
                            st.markdown("**Rubric**")
                            for r in rubric:
                                desc_r = r.get("description") or ""
                                st.write(f"- {desc_r}")
                                ratings = r.get("ratings") or []
                                if ratings:
                                    st.write("  Ratings:")
                                    for rt in ratings:
                                        st.write(f"    - {rt.get('description')} ({rt.get('points')})")

            with cols[1]:
                st.markdown("**Raw parsed JSON**")
                st.write("You can download the structured assignments/rubrics JSON below.")
                st.divider()
                parsed_json = json.dumps(parsed, indent=2)
                st.download_button(
                    "Download parsed JSON",
                    data=parsed_json,
                    file_name=f"{(st.session_state.docx_filename or 'assessor_guide').rsplit('.',1)[0]}-parsed.json",
                    mime="application/json",
                )
                # rubrics-only export (assignments/rubric compact)
                try:
                    compact = [
                        {"title": a.get("title"), "rubric": a.get("rubric", [])} for a in assignments
                    ]
                except Exception:
                    compact = parsed.get("assignments", [])
                rubrics_json = json.dumps(compact, indent=2)
                st.download_button(
                    "Download rubrics JSON",
                    data=rubrics_json,
                    file_name=f"{(st.session_state.docx_filename or 'assessor_guide').rsplit('.',1)[0]}-rubrics.json",
                    mime="application/json",
                )
