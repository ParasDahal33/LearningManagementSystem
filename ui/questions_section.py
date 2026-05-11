"""
ui/questions_section.py
Question editor UI and pagination.
"""
from __future__ import annotations

import math
import re
import streamlit as st

from core.utils import dedupe_questions, strip_q_prefix


def render_questions_section() -> None:
    st.divider()
    st.subheader("3) Questions")

    page_size = int(st.session_state.get("questions_page_size") or 10)
    if page_size not in {5, 10, 15, 20, 30}:
        page_size = 10
    questions = st.session_state.questions or []
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
                ak_e = q.get("assessor_key") or q.get("neutral_comments") or ""
                ak_edit_e = st.text_area("Correct answer / assessor comments (optional)", value=ak_e, key=f"{run}_assessor_essay_{i}", height=120)
                q["assessor_key"] = ak_edit_e.strip() or None

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
                ak = q.get("assessor_key") or ""
                ak_edit = st.text_area("Correct answer / assessor comments (optional)", value=ak, key=f"{run}_assessor_{i}", height=80)
                q["assessor_key"] = ak_edit.strip() or None

            else:
                opts = q.get("options", []) or []
                correct_set = set(q.get("correct", []) or [])
                st.write("**Options** (tick ✅ for correct answer)")
                new_opts: list[str] = []
                new_correct: list[int] = []
                for j, opt in enumerate(opts):
                    oc1, oc2 = st.columns([0.12, 0.88])
                    is_corr = oc1.checkbox(
                        f"Correct option {j + 1} for question {i + 1}",
                        value=(j in correct_set),
                        key=f"{run}_q{i}_corr_{j}",
                        label_visibility="hidden",
                    )
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
