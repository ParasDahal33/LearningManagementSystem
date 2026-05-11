"""
ui/sidebar.py
Sidebar rendering and session management for the Streamlit app.
"""
from __future__ import annotations

import re
import streamlit as st
from services.canvas_api import canvas_whoami, list_courses


def render_sidebar() -> str:
    """Render the sidebar and return the selected parser mode string.

    This function mutates `st.session_state` in place (login, course selection,
    parser mode, AI provider and keys). It returns the parser_mode value that
    the rest of the app should use.
    """
    login_expanded = not bool(st.session_state.logged_in)
    course_expanded = bool(st.session_state.logged_in) and not bool(st.session_state.selected_course_id)

    # --- Login ---
    with st.expander("🔐 Login", expanded=login_expanded):
        st.session_state.canvas_base_url = st.text_input(
            "Canvas Base URL", value="https://learningvault.instructure.com/"
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
        # Assessor guide import toggle
        st.session_state.assessor_import = st.checkbox("Assessor Guide import (structured)", value=bool(st.session_state.get("assessor_import", False)))
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

    return parser_mode
