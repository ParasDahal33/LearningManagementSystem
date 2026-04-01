"""
app.py
Canvas Quiz Uploader — main Streamlit entry point.
All business logic lives in core/, parsers/, and services/.
Run with:  streamlit run app.py
"""

from __future__ import annotations

import contextlib
import io
import math
import os
import re
import tempfile

import streamlit as st

# ---------------------------------------------------------------------------
# Internal modules
# ---------------------------------------------------------------------------
from core.config import init_session_state
from core.utils import (
    dedupe_questions,
    strip_q_prefix,
    v2_split_items_on_internal_qnums,
    v2c_clean_text,
    v2c_dedupe_questions,
    v2c_collapse_duplicate_mcq,
    v3_split_items_on_internal_qnums,
    v3_dedupe_questions,
)
from parsers.docx_extractor import (
    extract_items_with_red_v1,
    build_description_v1,
    v2_extract_items_with_red,
    v3_extract_items_with_red,
)
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
from services.canvas_api import (
    canvas_whoami,
    list_courses,
    get_existing_quiz_titles,
    generate_unique_title,
    create_canvas_quiz,
    publish_quiz,
    add_question_to_quiz,
    validate_before_upload,
)
from services.openai_services import OpenAIConfig, v3_ai_segment_items_openai
from services.gemini_services import GeminiConfig, v3_ai_segment_items_gemini

# ---------------------------------------------------------------------------
# Bootstrap session state
# ---------------------------------------------------------------------------
init_session_state()

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Canvas Quiz Uploader", layout="wide")
st.title("Canvas Quiz Uploader")

# ===========================================================================
# SIDEBAR
# ===========================================================================
with st.sidebar:
    login_expanded = not bool(st.session_state.logged_in)
    course_expanded = bool(st.session_state.logged_in) and not bool(st.session_state.selected_course_id)

    # --- Login ---
    with st.expander("🔐 Login", expanded=login_expanded):
        st.session_state.canvas_base_url = st.text_input(
            "Canvas Base URL", value=st.session_state.canvas_base_url
        ).strip()
        st.session_state.canvas_token = st.text_input(
            "Canvas Access Token", value=st.session_state.canvas_token, type="password"
        )
        c_login, c_logout = st.columns(2)
        if c_login.button("Login", use_container_width=True):
            try:
                me = canvas_whoami(
                    st.session_state.canvas_base_url, st.session_state.canvas_token
                )
                if me:
                    st.session_state.logged_in = True
                    st.session_state.me = me
                    st.session_state.courses_cache = None
                else:
                    st.session_state.logged_in = False
                    st.session_state.me = None
                    st.error("Login failed: token invalid/expired.")
            except Exception as e:
                st.session_state.logged_in = False
                st.session_state.me = None
                st.error(f"Login failed: {e}")

        if c_logout.button("Logout", use_container_width=True):
            for key in ("logged_in", "me", "selected_course_id", "courses_cache", "questions", "parsed_ok"):
                st.session_state[key] = (
                    False if key in ("logged_in", "parsed_ok") else None if key != "questions" else []
                )

        if st.session_state.logged_in and st.session_state.me:
            st.caption(f"User: {st.session_state.me.get('name', '')}")
        else:
            st.caption("Token login only.")

    # --- Course selector ---
    with st.expander("✅ Course", expanded=course_expanded):
        if not st.session_state.logged_in:
            st.info("Login first to load courses.")
        else:
            try:
                if st.session_state.courses_cache is None:
                    st.session_state.courses_cache = list_courses(
                        st.session_state.canvas_base_url, st.session_state.canvas_token
                    )
                courses = st.session_state.courses_cache or []
                if not courses:
                    st.warning("No courses visible to this token.")
                else:
                    label_to_id: dict[str, str] = {}
                    labels: list[str] = []
                    for c in courses:
                        cid = c.get("id")
                        name = (c.get("name") or c.get("course_code") or f"Course {cid}").strip()
                        label = f"{name} (ID: {cid})"
                        labels.append(label)
                        label_to_id[label] = str(cid)

                    default_index = 0
                    if st.session_state.selected_course_id:
                        for i, lb in enumerate(labels):
                            if label_to_id[lb] == st.session_state.selected_course_id:
                                default_index = i
                                break
                    chosen = st.selectbox("Select course", labels, index=default_index)
                    st.session_state.selected_course_id = label_to_id[chosen]

                    if st.button("Refresh courses", use_container_width=True):
                        st.session_state.courses_cache = None
                        st.rerun()
            except Exception as e:
                st.error(f"Failed to load courses: {e}")

    # --- Parser selector ---
    with st.expander("Parser", expanded=True):
        parser_mode = st.selectbox(
            "Version",
            ["v1 (rule-based)", "v2 (rule-based)", "v3 (AI-hybrid)"],
            index=0,
        )
        prev_mode = st.session_state.last_parser_mode
        if prev_mode is None:
            st.session_state.last_parser_mode = parser_mode
        elif prev_mode != parser_mode:
            st.session_state.last_parser_mode = parser_mode
            st.session_state.questions = []
            st.session_state.parsed_ok = False
            st.session_state.description_html = ""
            st.session_state.docx_filename = None
            st.session_state.parse_run_id += 1
            st.rerun()

        if parser_mode.startswith("v3"):
            st.divider()
            st.session_state.ai_provider = st.radio(
                "AI Provider", ["OpenAI", "Gemini"], index=0 if st.session_state.get("ai_provider") != "Gemini" else 1
            )
            if st.session_state.ai_provider == "OpenAI":
                st.session_state.openai_api_key = st.text_input(
                    "OpenAI API key", value=st.session_state.openai_api_key, type="password"
                )
                st.session_state.openai_model = st.text_input(
                    "Model", value=st.session_state.openai_model
                )
                st.session_state.openai_base_url = st.text_input(
                    "Base URL", value=st.session_state.openai_base_url
                )
            else:
                st.session_state.gemini_api_key = st.text_input(
                    "Gemini API key", value=st.session_state.get("gemini_api_key", ""), type="password"
                )
                st.session_state.gemini_model = st.text_input(
                    "Model", value=st.session_state.get("gemini_model", "gemini-1.5-flash")
                )
                st.session_state.gemini_base_url = st.text_input(
                    "Base URL", value=st.session_state.get("gemini_base_url", "https://generativelanguage.googleapis.com")
                )

# ===========================================================================
# AUTH GATE
# ===========================================================================
if not st.session_state.logged_in:
    st.warning("Please login in the sidebar first.")
    st.stop()

if not st.session_state.selected_course_id:
    st.warning("Please select a course in the sidebar.")
    st.stop()

course_id = st.session_state.selected_course_id
canvas_base_url = st.session_state.canvas_base_url
canvas_token = st.session_state.canvas_token

# ===========================================================================
# SECTION 1 — UPLOAD & PARSE
# ===========================================================================
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
                    model=(st.session_state.openai_model or "gpt-4.1-mini").strip(),
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

questions = st.session_state.questions or []
if not questions:
    st.stop()

# ===========================================================================
# SECTION 2 — QUIZ SETTINGS
# ===========================================================================
st.subheader("2) Details (Canvas Quiz Settings)")
default_title = os.path.splitext(st.session_state.docx_filename or "Quiz")[0]
d = st.session_state.details
run = st.session_state.parse_run_id

quiz_title = st.text_input(
    "Quiz Title *",
    value=(d.get("quiz_title") or default_title),
    key=f"{run}_quiz_title",
)
quiz_instructions = st.text_area(
    "Quiz Instructions (HTML allowed)",
    value=(d.get("quiz_instructions") or st.session_state.description_html or ""),
    height=180,
    key=f"{run}_quiz_instructions",
)
d["quiz_title"] = quiz_title
d["quiz_instructions"] = quiz_instructions

c1, c2, c3 = st.columns(3)
d["shuffle_answers"] = c1.checkbox("Shuffle Answers", value=bool(d.get("shuffle_answers", True)))
d["one_question_at_a_time"] = c2.checkbox("Show one question at a time", value=bool(d.get("one_question_at_a_time", False)))
d["show_correct_answers"] = c3.checkbox("Let Students See The Correct Answers", value=bool(d.get("show_correct_answers", False)))

c4, c5, c6 = st.columns(3)
d["time_limit"] = c4.number_input("Time Limit (minutes, 0 = none)", min_value=0, max_value=1440, value=int(d.get("time_limit", 0) or 0), step=5)
d["allow_multiple_attempts"] = c5.checkbox("Allow Multiple Attempts", value=bool(d.get("allow_multiple_attempts", False)))
if d["allow_multiple_attempts"]:
    cur_attempts = max(2, int(d.get("allowed_attempts", 2) or 2))
    d["allowed_attempts"] = c5.number_input("Allowed Attempts", min_value=2, max_value=20, value=cur_attempts, step=1)
else:
    d["allowed_attempts"] = 1
d["scoring_policy"] = c6.selectbox(
    "Quiz Score to Keep",
    ["keep_highest", "keep_latest"],
    index=0 if d.get("scoring_policy", "keep_highest") == "keep_highest" else 1,
)

st.markdown("**Quiz Restrictions**")
d["access_code_enabled"] = st.checkbox("Require an access code", value=bool(d.get("access_code_enabled", False)))
if d["access_code_enabled"]:
    d["access_code"] = st.text_input("Access code", value=(d.get("access_code") or ""))
else:
    d["access_code"] = ""

st.markdown("**Assign / Availability (optional, ISO datetime)**")
st.caption("Example: 2026-01-20T23:59:00Z  (leave blank if unsure)")
cc1, cc2, cc3 = st.columns(3)
d["due_at"] = cc1.text_input("Due Date (due_at)", value=(d.get("due_at") or ""))
d["unlock_at"] = cc2.text_input("Available from (unlock_at)", value=(d.get("unlock_at") or ""))
d["lock_at"] = cc3.text_input("Until (lock_at)", value=(d.get("lock_at") or ""))
st.session_state.details = d

# ===========================================================================
# SECTION 3 — QUESTION EDITOR
# ===========================================================================
st.divider()
st.subheader("3) Questions")

page_size = int(st.session_state.get("questions_page_size") or 10)
if page_size not in {5, 10, 15, 20, 30}:
    page_size = 10
total = len(questions)
total_pages = max(1, math.ceil(total / page_size))
page = max(1, min(int(st.session_state.get("questions_page") or 1), total_pages))

start = (page - 1) * page_size
end = min(start + page_size, total)
st.caption(f"Showing questions {start + 1}–{end} of {total}")

edited = [q.copy() for q in questions]
run = st.session_state.parse_run_id

for i in range(start, end):
    q = edited[i]
    kind = (q.get("kind") or "").lower()
    preview = strip_q_prefix(q.get("question", ""))[:90]
    label_kind = "Matching" if kind == "matching" else ("Essay/Short Answer" if kind == "essay" else "MCQ")

    with st.expander(f"Q{i + 1} ({label_kind}): {preview}"):
        q_text = st.text_area(
            "Question text", value=q.get("question", ""), key=f"{run}_qtext_{i}", height=90
        )
        q["question"] = strip_q_prefix(q_text.strip())

        if kind == "essay":
            st.info("This question will be uploaded as an ESSAY (student types the answer).")
            q["options"] = []
            q["correct"] = []
            q["multi"] = False

        elif kind == "matching":
            st.info("This question will be uploaded as MATCHING (left item → dropdown right item).")
            st.caption("Tip: if right side was a bullet list in Word, it appears joined with '; ' — that is correct.")
            pairs = q.get("pairs") or []
            new_pairs = []
            for j, p in enumerate(pairs):
                lc1, lc2 = st.columns([0.6, 0.4])
                left = lc1.text_input(f"Left (row {j + 1})", value=p.get("left", ""), key=f"{run}_match_{i}_l_{j}")
                right = lc2.text_input(f"Right (row {j + 1})", value=p.get("right", ""), key=f"{run}_match_{i}_r_{j}")
                if left.strip() and right.strip():
                    new_pairs.append({"left": left.strip(), "right": right.strip()})
            q["pairs"] = new_pairs

        else:
            opts = q.get("options", []) or []
            correct_set = set(q.get("correct", []) or [])
            st.write("**Options** (tick ✅ for correct answer)")
            new_opts: list[str] = []
            new_correct: list[int] = []
            for j, opt in enumerate(opts):
                oc1, oc2 = st.columns([0.12, 0.88])
                is_corr = oc1.checkbox("", value=(j in correct_set), key=f"{run}_q{i}_corr_{j}")
                opt_text = oc2.text_input(f"Option {j + 1}", value=opt, key=f"{run}_q{i}_opt_{j}")
                new_opts.append(opt_text.strip())
                if is_corr:
                    new_correct.append(j)

            add_opt = st.text_input("New option text (optional)", value="", key=f"{run}_q{i}_newopt")
            if add_opt.strip():
                new_opts.append(add_opt.strip())

            cleaned_opts: list[str] = []
            idx_map: dict[int, int] = {}
            for old_index, txt in enumerate(new_opts):
                if txt.strip():
                    idx_map[old_index] = len(cleaned_opts)
                    cleaned_opts.append(txt.strip())

            remapped_correct = [idx_map[old_i] for old_i in new_correct if old_i in idx_map]
            q["options"] = cleaned_opts
            q["correct"] = sorted(set(remapped_correct))
            qlower = (q.get("question") or "").lower()
            q["multi"] = (
                "apply" in qlower
                or len(q["correct"]) > 1
                or bool(re.search(r"\bselect\s+(two|three|four|five|\d+)", qlower))
            )

edited = dedupe_questions(edited)
st.session_state.questions = edited

# Pagination controls
st.divider()
pc1, pc2 = st.columns([0.55, 0.45])
new_page_size = pc1.selectbox("Questions per page", [5, 10, 15, 20, 30], index=[5, 10, 15, 20, 30].index(page_size))
new_total_pages = max(1, math.ceil(total / int(new_page_size)))
new_page = (
    pc2.selectbox("Page", list(range(1, new_total_pages + 1)), index=max(0, min(page, new_total_pages) - 1), format_func=lambda n: f"{n} / {new_total_pages}")
    if new_total_pages <= 200
    else pc2.number_input("Page", min_value=1, max_value=new_total_pages, value=min(page, new_total_pages), step=1)
)
st.session_state.questions_page_size = int(new_page_size)
st.session_state.questions_page = int(new_page)

# ===========================================================================
# SECTION 4 — SAVE TO CANVAS
# ===========================================================================
st.divider()
st.subheader("4) Save to Canvas")
colS1, colS2 = st.columns([1, 1])
save_draft = colS1.button("💾 Save to Canvas (Draft)")
save_publish = colS2.button("🚀 Save & Publish")

if save_draft or save_publish:
    qs = st.session_state.questions or []
    probs = validate_before_upload(qs)
    if probs:
        st.error("Please fix these issues before uploading:")
        for p in probs[:15]:
            st.write(f"- {p}")
        st.stop()

    base_title = (quiz_title or "").strip() or default_title
    try:
        existing_titles = get_existing_quiz_titles(canvas_base_url, course_id, canvas_token)
        final_title = generate_unique_title(base_title, existing_titles)

        with st.spinner("Creating quiz in Canvas..."):
            quiz_id = create_canvas_quiz(
                canvas_base_url=canvas_base_url,
                course_id=course_id,
                canvas_token=canvas_token,
                title=final_title,
                description_html=quiz_instructions,
                settings=st.session_state.details,
            )

        with st.spinner("Uploading questions..."):
            for q in qs:
                add_question_to_quiz(canvas_base_url, course_id, canvas_token, quiz_id, q)

        if save_publish:
            with st.spinner("Publishing quiz..."):
                publish_quiz(canvas_base_url, course_id, canvas_token, quiz_id)

        st.success("✅ Done!")
        st.write(f"**Quiz title:** {final_title}")
        st.write(f"**Quiz ID:** {quiz_id}")
        st.write(f"**Course ID:** {course_id}")
        st.info("Quiz published ✅" if save_publish else "Quiz saved as draft (unpublished).")
    except Exception as e:
        st.error(f"❌ Upload failed: {e}")

st.caption("Token login only (Canvas API does not support username/password).")