"""
services/canvas_api.py
Canvas LMS REST API client functions:
authentication, course listing, quiz creation, question upload, publishing.
"""

from __future__ import annotations

import re

import requests

from .util import strip_q_prefix


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _canvas_headers(canvas_token: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {canvas_token}"}


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
def canvas_whoami(canvas_base_url: str, canvas_token: str):
    """Return the current user object or None if the token is invalid."""
    url = f"{canvas_base_url.rstrip('/')}/users/self"
    r = requests.get(url, headers=_canvas_headers(canvas_token), timeout=30)
    if r.status_code == 401:
        return None
    r.raise_for_status()
    return r.json()


# ---------------------------------------------------------------------------
# Courses
# ---------------------------------------------------------------------------
def list_courses(canvas_base_url: str, canvas_token: str) -> list[dict]:
    """Return all courses visible to the token (all pages)."""
    url = f"{canvas_base_url.rstrip('/')}/courses"
    out: list[dict] = []
    page = 1
    while True:
        r = requests.get(
            url,
            headers=_canvas_headers(canvas_token),
            params={"per_page": 100, "page": page},
            timeout=60,
        )
        r.raise_for_status()
        batch = r.json()
        if not batch:
            break
        out.extend(batch)
        page += 1
    return out


# ---------------------------------------------------------------------------
# Quizzes
# ---------------------------------------------------------------------------
def get_existing_quiz_titles(canvas_base_url: str, course_id: str, canvas_token: str) -> set[str]:
    """Return the set of existing quiz titles for the course."""
    url = f"{canvas_base_url.rstrip('/')}/courses/{course_id}/quizzes"
    titles: set[str] = set()
    page = 1
    while True:
        r = requests.get(
            url,
            headers=_canvas_headers(canvas_token),
            params={"page": page, "per_page": 100},
            timeout=60,
        )
        r.raise_for_status()
        data = r.json()
        if not data:
            break
        for q in data:
            titles.add((q.get("title") or "").strip())
        page += 1
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
    settings = settings or {}
    url = f"{canvas_base_url.rstrip('/')}/courses/{course_id}/quizzes"

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

    r = requests.post(
        url,
        headers=_canvas_headers(canvas_token),
        json={"quiz": quiz_obj},
        timeout=60,
    )
    if r.status_code == 401:
        raise RuntimeError("401 Unauthorized — token invalid/expired.")
    if r.status_code == 403:
        raise RuntimeError("403 Forbidden — missing permission in this course.")
    r.raise_for_status()
    return r.json()["id"]


def publish_quiz(canvas_base_url: str, course_id: str, canvas_token: str, quiz_id: int) -> None:
    """Publish (make visible to students) an existing quiz."""
    url = f"{canvas_base_url.rstrip('/')}/courses/{course_id}/quizzes/{quiz_id}"
    r = requests.put(
        url,
        headers=_canvas_headers(canvas_token),
        json={"quiz": {"published": True}},
        timeout=60,
    )
    r.raise_for_status()


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
    url = f"{canvas_base_url.rstrip('/')}/courses/{course_id}/quizzes/{quiz_id}/questions"
    qtext = strip_q_prefix((q.get("question") or "").strip())
    kind = (q.get("kind") or "").lower()

    # --- Matching ---
    if kind == "matching":
        answers = [
            {"answer_match_left": p.get("left", "").strip(), "answer_match_right": p.get("right", "").strip(), "answer_weight": 100}
            for p in (q.get("pairs") or [])
            if p.get("left", "").strip() and p.get("right", "").strip()
        ]
        payload = {
            "question": {
                "question_name": (qtext[:100] if qtext else "Matching"),
                "question_text": qtext,
                "question_type": "matching_question",
                "points_possible": 1,
                "answers": answers,
            }
        }
        r = requests.post(url, headers=_canvas_headers(canvas_token), json=payload, timeout=60)
        if r.status_code >= 400:
            raise RuntimeError(f"Canvas error {r.status_code}: {r.text[:600]}")
        r.raise_for_status()
        return

    opts = [o.strip() for o in (q.get("options") or []) if o and o.strip()]
    correct = q.get("correct", []) or []

    # --- Essay / Short Answer ---
    if kind == "essay" or len(opts) < 2:
        payload = {
            "question": {
                "question_name": (qtext[:100] if qtext else "Question"),
                "question_text": qtext or " ",
                "question_type": "essay_question",
                "points_possible": 1,
            }
        }
        r = requests.post(url, headers=_canvas_headers(canvas_token), json=payload, timeout=60)
        r.raise_for_status()
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
    payload = {
        "question": {
            "question_name": (qtext[:100] if qtext else "Question"),
            "question_text": qtext,
            "question_type": qtype,
            "points_possible": 1,
            "answers": answers,
        }
    }
    r = requests.post(url, headers=_canvas_headers(canvas_token), json=payload, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(f"Canvas error {r.status_code}: {r.text[:600]}")
    r.raise_for_status()


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