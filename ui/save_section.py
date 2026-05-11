"""
ui/save_section.py
Handles saving the quiz and questions to Canvas.
"""
from __future__ import annotations

import streamlit as st

from services.canvas_api import (
    get_existing_quiz_titles,
    generate_unique_title,
    create_canvas_quiz,
    publish_quiz,
    add_question_to_quiz,
    validate_before_upload,
)


def render_save_section(canvas_base_url: str, course_id: str) -> None:
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

        default_title = st.session_state.details.get("quiz_title") or "Quiz"
        base_title = (default_title or "").strip() or default_title
        try:
            existing_titles = get_existing_quiz_titles(canvas_base_url, course_id, st.session_state.canvas_token)
            final_title = generate_unique_title(base_title, existing_titles)

            with st.spinner("Creating quiz in Canvas..."):
                quiz_id = create_canvas_quiz(
                    canvas_base_url=canvas_base_url,
                    course_id=course_id,
                    canvas_token=st.session_state.canvas_token,
                    title=final_title,
                    description_html=st.session_state.details.get("quiz_instructions"),
                    settings=st.session_state.details,
                )

            with st.spinner("Uploading questions..."):
                for q in qs:
                    add_question_to_quiz(canvas_base_url, course_id, st.session_state.canvas_token, quiz_id, q)

            if save_publish:
                with st.spinner("Publishing quiz..."):
                    publish_quiz(canvas_base_url, course_id, st.session_state.canvas_token, quiz_id)

            st.success("✅ Done!")
            st.write(f"**Quiz title:** {final_title}")
            st.write(f"**Quiz ID:** {quiz_id}")
            st.write(f"**Course ID:** {course_id}")
            st.info("Quiz published ✅" if save_publish else "Quiz saved as draft (unpublished).")
        except Exception as e:
            st.error(f"❌ Upload failed: {e}")

    st.caption("Token login only (Canvas API does not support username/password).")
