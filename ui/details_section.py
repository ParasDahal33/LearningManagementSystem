"""
ui/details_section.py
Render and manage quiz details/settings UI.
"""
from __future__ import annotations

import os
import streamlit as st


def render_details_section() -> None:
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
