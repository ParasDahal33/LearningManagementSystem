"""
ui/rubrics_section.py
Editable UI for parsed Assessor Guide rubrics and assignments.
"""
from __future__ import annotations

import json
import streamlit as st

from services.canvas_api import (
    create_canvas_assignment,
    create_canvas_rubric,
    validate_rubrics_before_upload,
)

def render_rubrics_section() -> None:
    if not st.session_state.get("assessor_parse"):
        return

    st.divider()
    st.subheader("1.1) Parsed Assessor Guide")
    
    parsed = st.session_state.get("assessor_parse")
    run = st.session_state.parse_run_id
    assignments = parsed.get("assignments", []) if isinstance(parsed, dict) else []

    if not assignments:
        st.info("No assignments found in parsed output.")
        return

    # Create an editable copy/view
    for i, assign in enumerate(assignments):
        title = assign.get("title") or f"Assignment {i + 1}"
        
        with st.expander(f"Assignment: {title}", expanded=True):
            # Assignment Title and Description
            new_title = st.text_input("Assignment Title", value=title, key=f"{run}_rub_title_{i}")
            assign["title"] = new_title

            desc = assign.get("description_html") or assign.get("description") or ""
            new_desc = st.text_area("Description (HTML)", value=desc, key=f"{run}_rub_desc_{i}", height=150)
            assign["description_html"] = new_desc

            sub_types = assign.get("submission_types") or []
            new_sub_types = st.text_input("Submission Types (comma separated)", value=", ".join(sub_types), key=f"{run}_rub_sub_{i}")
            assign["submission_types"] = [s.strip() for s in new_sub_types.split(",") if s.strip()]

            # Rubric Criteria
            rubric = assign.get("rubric") or []
            if rubric:
                st.markdown("**Rubric Criteria**")
                for j, crit in enumerate(rubric):
                    c_desc = crit.get("description") or ""
                    new_c_desc = st.text_input(f"Criterion {j + 1}", value=c_desc, key=f"{run}_rub_{i}_crit_{j}")
                    crit["description"] = new_c_desc

                    # Ratings (Satisfactory / Not Yet Satisfactory)
                    ratings = crit.get("ratings") or []
                    cols = st.columns(len(ratings) if ratings else 1)
                    for k, rate in enumerate(ratings):
                        with cols[k]:
                            r_label = rate.get("description") or ""
                            r_pts = rate.get("points") or 0.0
                            new_label = st.text_input(f"Rating {k + 1} label", value=r_label, key=f"{run}_rub_{i}_{j}_rl_{k}")
                            new_pts = st.number_input(f"Points", value=float(r_pts), key=f"{run}_rub_{i}_{j}_rp_{k}", step=0.5)
                            rate["description"] = new_label
                            rate["points"] = new_pts

    # Update session state
    st.session_state.assessor_parse["assignments"] = assignments

    # --- Save to Canvas Section ---
    st.divider()
    st.subheader("1.2) Save Assessor Guide to Canvas")
    st.info("This will create one Canvas Assignment and one Rubric for each stage listed above.")
    
    col_save1, col_save2 = st.columns([1, 1])
    save_btn = col_save1.button("💾 Save Assignments & Rubrics to Canvas", type="primary", use_container_width=True)

    if save_btn:
        problems = validate_rubrics_before_upload(assignments)
        if problems:
            st.error("Please fix these issues before uploading:")
            for p in problems:
                st.write(f"- {p}")
        else:
            course_id = st.session_state.selected_course_id
            canvas_url = st.session_state.canvas_base_url
            token = st.session_state.canvas_token
            
            success_count = 0
            progress_bar = st.progress(0)
            
            try:
                for idx, assign in enumerate(assignments):
                    with st.status(f"Uploading {assign['title']}...", expanded=False) as status:
                        # 1. Create Assignment
                        st.write("Creating assignment...")
                        a_id = create_canvas_assignment(
                            canvas_url, course_id, token,
                            name=assign["title"],
                            description_html=assign["description_html"],
                            submission_types=assign["submission_types"]
                        )
                        
                        # 2. Create and Associate Rubric
                        if assign.get("rubric"):
                            rubric_data = assign["rubric"]
                            st.write(f"Creating rubric with {len(rubric_data)} criteria...")
                            for c_idx, crit in enumerate(rubric_data):
                                st.write(f"  - **Criterion {c_idx+1}**: {crit.get('description', 'No description')}")
                                for r_idx, rate in enumerate(crit.get("ratings", [])):
                                    st.write(f"    * Rating: {rate.get('description')} ({rate.get('points')} pts)")
                            
                            create_canvas_rubric(
                                canvas_url, course_id, token,
                                assignment_id=a_id,
                                title=f"Rubric: {assign['title']}",
                                criteria=rubric_data
                            )
                        status.update(label=f"✅ {assign['title']} Uploaded", state="complete")
                        success_count += 1
                        progress_bar.progress((idx + 1) / len(assignments))
                st.success(f"Successfully created {success_count} assignments with rubrics!")
            except Exception as e:
                st.error(f"Upload failed: {e}")

    # Export options
    st.divider()
    c1, c2 = st.columns(2)
    filename_base = (st.session_state.get("docx_filename") or "assessor_guide").rsplit('.', 1)[0]
    
    with c1:
        full_json = json.dumps(st.session_state.assessor_parse, indent=2)
        st.download_button("Download Full JSON", data=full_json, file_name=f"{filename_base}-parsed.json", mime="application/json")
    
    with c2:
        compact = [{"title": a.get("title"), "rubric": a.get("rubric", [])} for a in assignments]
        compact_json = json.dumps(compact, indent=2)
        st.download_button("Download Rubrics Only", data=compact_json, file_name=f"{filename_base}-rubrics.json", mime="application/json")