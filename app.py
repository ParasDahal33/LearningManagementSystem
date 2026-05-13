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
import json

# ---------------------------------------------------------------------------
# Internal modules
# ---------------------------------------------------------------------------
from core.config import init_session_state
from ui.sidebar import render_sidebar
from ui.parse_section import render_parse_section
from ui.details_section import render_details_section
from ui.rubrics_section import render_rubrics_section
from ui.questions_section import render_questions_section
from ui.save_section import render_save_section

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
# render sidebar and return the parser mode selected
with st.sidebar:
    parser_mode = render_sidebar()

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

# Render parse section (handles upload/parse and assessor guide preview)
render_parse_section(parser_mode)

if st.session_state.get("parsed_ok"):
    # Render rubrics section if assessor data exists
    if st.session_state.get("assessor_parse"):
        render_rubrics_section()

    # Stop early if there are no quiz questions (prevents rendering settings/questions for Rubrics mode)
    questions = st.session_state.questions or []
    if not questions:
        st.stop()


# Render remaining sections
render_details_section()
render_questions_section()
render_save_section(canvas_base_url, course_id)