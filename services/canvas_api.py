"""
services/canvas_api.py
Canvas LMS REST API client functions:
authentication, course listing, quiz creation, question upload, publishing.
"""

from __future__ import annotations

import re

from canvasapi import Canvas
from canvasapi.exceptions import CanvasException

from core.utils import strip_q_prefix


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
def canvas_whoami(canvas_base_url: str, canvas_token: str):
    """Return the current user object or None if the token is invalid."""
    try:
        canvas = Canvas(canvas_base_url, canvas_token)
        user = canvas.get_current_user()
        return {
            "id": user.id,
            "name": user.name,
            "login_id": getattr(user, "login_id", getattr(user, "email", "")),
        }
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Courses
# ---------------------------------------------------------------------------
def get_course(canvas_base_url: str, canvas_token: str, course_id: str) -> dict | None:
    """
    Retrieve a specific course by ID to verify existence.
    """
    try:
        canvas = Canvas(canvas_base_url, canvas_token)
        course = canvas.get_course(course_id)
        # Return minimal dict to satisfy potential callers expecting a dict
        return {"id": course.id, "name": getattr(course, "name", "")}
    except Exception:
        return None


def list_courses(canvas_base_url: str, canvas_token: str) -> list[dict]:
    """Return all courses visible to the token (all pages)."""
    canvas = Canvas(canvas_base_url, canvas_token)
    out: list[dict] = []
    try:
        courses = canvas.get_courses(per_page=100)
        for c in courses:
            out.append(
                {
                    "id": c.id,
                    "name": (
                        getattr(c, "name", "")
                        or getattr(c, "course_code", "")
                        or f"Course {c.id}"
                    ).strip(),
                    "course_code": getattr(c, "course_code", ""),
                }
            )
    except Exception:
        # Propagate errors (e.g. auth failure) to be caught by the caller
        raise
    return out


# ---------------------------------------------------------------------------
# Quizzes
# ---------------------------------------------------------------------------
def get_existing_quiz_titles(canvas_base_url: str, course_id: str, canvas_token: str) -> set[str]:
    """Return the set of existing quiz titles for the course."""
    canvas = Canvas(canvas_base_url, canvas_token)
    course = canvas.get_course(course_id)
    titles: set[str] = set()
    # canvasapi handles pagination automatically
    for q in course.get_quizzes(per_page=100):
        t = getattr(q, "title", "")
        if t:
            titles.add(t.strip())
    return titles


def generate_unique_title(base_title: str, existing_titles: set[str]) -> str:
    """Append an incrementing counter if base_title already exists."""
    if base_title not in existing_titles:
        return base_title
    i = 1
    while True:
        candidate = f"{base_title} ({i})"
        if candidate not in existing_titles:
            return candidate
        i += 1


def create_canvas_quiz(
    canvas_base_url: str,
    course_id: str,
    canvas_token: str,
    *,
    title: str,
    description_html: str = "",
    settings: dict | None = None,
) -> int:
    """Create a new quiz and return its quiz_id."""
    canvas = Canvas(canvas_base_url, canvas_token)
    course = canvas.get_course(course_id)
    settings = settings or {}

    quiz_obj: dict = {
        "title": title,
        "description": description_html,
        "published": False,
        "quiz_type": "assignment",
        "shuffle_answers": bool(settings.get("shuffle_answers", False)),
        "one_question_at_a_time": bool(settings.get("one_question_at_a_time", False)),
        "show_correct_answers": bool(settings.get("show_correct_answers", False)),
        "scoring_policy": settings.get("scoring_policy", "keep_highest"),
    }

    tl = int(settings.get("time_limit", 0) or 0)
    if tl > 0:
        quiz_obj["time_limit"] = tl

    if bool(settings.get("allow_multiple_attempts", False)):
        aa = max(2, int(settings.get("allowed_attempts", 2) or 2))
        quiz_obj["allowed_attempts"] = aa

    if bool(settings.get("access_code_enabled", False)) and (settings.get("access_code") or "").strip():
        quiz_obj["access_code"] = settings["access_code"].strip()

    for k in ["due_at", "unlock_at", "lock_at"]:
        v = (settings.get(k) or "").strip()
        if v:
            quiz_obj[k] = v

    try:
        quiz = course.create_quiz(quiz_obj)
        return quiz.id
    except CanvasException as e:
        raise RuntimeError(f"Canvas API Error: {e}")


def publish_quiz(canvas_base_url: str, course_id: str, canvas_token: str, quiz_id: int) -> None:
    """Publish (make visible to students) an existing quiz."""
    canvas = Canvas(canvas_base_url, canvas_token)
    course = canvas.get_course(course_id)
    quiz = course.get_quiz(quiz_id)
    quiz.edit(quiz={"published": True})


# ---------------------------------------------------------------------------
# Questions
# ---------------------------------------------------------------------------
def add_question_to_quiz(
    canvas_base_url: str,
    course_id: str,
    canvas_token: str,
    quiz_id: int,
    q: dict,
) -> None:
    """Upload a single question dict to an existing quiz."""
    canvas = Canvas(canvas_base_url, canvas_token)
    course = canvas.get_course(course_id)
    quiz = course.get_quiz(quiz_id)

    qtext = strip_q_prefix((q.get("question") or "").strip())
    kind = (q.get("kind") or "").lower()

    # --- Matching ---
    if kind == "matching":
        answers = [
            {"answer_match_left": p.get("left", "").strip(), "answer_match_right": p.get("right", "").strip(), "answer_weight": 100}
            for p in (q.get("pairs") or [])
            if p.get("left", "").strip() and p.get("right", "").strip()
        ]
        q_params = {
            "question_name": (qtext[:100] if qtext else "Matching"),
            "question_text": qtext,
            "question_type": "matching_question",
            "points_possible": 1,
            "answers": answers,
        }
        quiz.create_question(question=q_params)
        return

    opts = [o.strip() for o in (q.get("options") or []) if o and o.strip()]
    correct = q.get("correct", []) or []

    # --- Essay / Short Answer ---
    if kind == "essay" or len(opts) < 2:
        q_params = {
            "question_name": (qtext[:100] if qtext else "Question"),
            "question_text": qtext or " ",
            "question_type": "essay_question",
            "points_possible": 1,
        }
        quiz.create_question(question=q_params)
        return

    # --- MCQ / Multiple Answers ---
    qlower = (qtext or "").lower()
    multi = (
        bool(q.get("multi"))
        or len(correct) > 1
        or bool(re.search(r"\bselect\s+(two|three|four|five|\d+)", qlower))
        or "apply" in qlower
    )
    qtype = "multiple_answers_question" if multi else "multiple_choice_question"
    answers = [
        {"answer_text": opt, "answer_weight": (100 if idx in correct else 0)}
        for idx, opt in enumerate(opts)
    ]
    q_params = {
        "question_name": (qtext[:100] if qtext else "Question"),
        "question_text": qtext,
        "question_type": qtype,
        "points_possible": 1,
        "answers": answers,
    }
    quiz.create_question(question=q_params)


# ---------------------------------------------------------------------------
# Upload validation
# ---------------------------------------------------------------------------
def validate_before_upload(qs: list[dict]) -> list[str]:
    """Return a list of human-readable problem strings (empty = OK to upload)."""
    problems: list[str] = []
    for idx, q in enumerate(qs, start=1):
        kind = (q.get("kind") or "").lower()
        qt = (q.get("question") or "").strip()
        if len(qt) < 10:
            problems.append(f"Q{idx}: question text too short.")
        if kind == "matching":
            if len(q.get("pairs") or []) < 2:
                problems.append(f"Q{idx}: matching needs at least 2 pairs.")
        elif kind != "essay":
            opts = q.get("options") or []
            if len(opts) >= 2 and not (q.get("correct") or []):
                problems.append(f"Q{idx}: no correct answer selected (red not detected or tick ✅).")
    return problems