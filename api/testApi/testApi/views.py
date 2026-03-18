"""
quiz/views.py
All API views for the Canvas Quiz Uploader.

Endpoint summary
----------------
POST /api/v1/auth/verify/          Verify a Canvas token and return user info
GET  /api/v1/courses/              List courses for a Canvas token
POST /api/v1/parse/                Upload a DOCX and parse questions
POST /api/v1/questions/validate/   Validate a list of questions before upload
POST /api/v1/quiz/upload/          Create a quiz + upload questions to Canvas
"""
from __future__ import annotations

import os
import re
import tempfile

from rest_framework import status
from rest_framework.parsers import MultiPartParser, JSONParser
from rest_framework.response import Response
from rest_framework.views import APIView

from .Serializers import (
    CanvasAuthSerializer,
    ParseRequestSerializer,
    UploadQuizRequestSerializer,
    ValidateQuestionsSerializer,
)

# Business-logic imports (unchanged from the original modular project)
from .util import (
    dedupe_questions,
    v2_split_items_on_internal_qnums,
    v2c_dedupe_questions,
    v2c_collapse_duplicate_mcq,
    v3_split_items_on_internal_qnums,
    v3_dedupe_questions,
)
from .docx_extractor import (
    extract_items_with_red_v1,
    build_description_v1,
    v3_extract_items_with_red,
)
from .mcq_parser import (
    parse_mcq_questions_v1,
    parse_essay_questions_v1,
    v2c_merge_dangling_question_lines,
    v2c_parse_mcq_questions,
    v2c_parse_essay_questions,
    v3_parse_essay_questions_rule_based,
    v3_filter_items_for_ai,
)
from .matching_parser import (
    parse_matching_questions_doc_order_v1_exact,
    v3_parse_matching_questions_doc_order,
    v3_parse_table_defined_terms_as_essays,
    v3_parse_table_characteristics_as_essays,
    v3_collect_ignore_texts_from_forced_tables,
)
from .canvas_api import (
    canvas_whoami,
    list_courses,
    get_existing_quiz_titles,
    generate_unique_title,
    create_canvas_quiz,
    publish_quiz,
    add_question_to_quiz,
    validate_before_upload,
)
from .openai_service import OpenAIConfig, v3_ai_segment_items_openai


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _run_parser(docx_path: str, parser_mode: str, openai_cfg: OpenAIConfig | None):
    """
    Run the appropriate parser and return (questions, description_html, debug_info).
    Mirrors the parse_btn logic from app.py exactly.
    """
    desc_items = extract_items_with_red_v1(docx_path)
    description_html = build_description_v1(desc_items)

    qs: list[dict] = []
    ai_log: list[str] = []
    removed_dupes = 0

    # ------------------------------------------------------------------ v1
    if parser_mode == "v1":
        matching = parse_matching_questions_doc_order_v1_exact(docx_path)
        mcq = parse_mcq_questions_v1(desc_items)
        essay = parse_essay_questions_v1(desc_items)
        qs = dedupe_questions(matching + mcq + essay)

    # ------------------------------------------------------------------ v2
    elif parser_mode == "v2":
        items_v2 = v2c_merge_dangling_question_lines(desc_items)
        items_v2 = v2_split_items_on_internal_qnums(items_v2)
        matching = parse_matching_questions_doc_order_v1_exact(docx_path)
        mcq = v2c_parse_mcq_questions(items_v2)
        essay = v2c_parse_essay_questions(items_v2)
        qs = v2c_dedupe_questions(matching + mcq + essay)
        qs = v2c_collapse_duplicate_mcq(qs)

    # ------------------------------------------------------------------ v3
    else:
        if openai_cfg is None:
            raise ValueError("v3 parser requires an OpenAI API key.")
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
        ai_input = v3_filter_items_for_ai(
            items, ignore_terms=ignore_terms, ignore_texts=ignore_texts, mode="balanced"
        )
        ai_qs, ai_log = v3_ai_segment_items_openai(ai_input, openai_cfg)
        rule_essays = v3_parse_essay_questions_rule_based(items)

        qs = matching + table_essays + ai_qs + rule_essays
        qs.sort(key=lambda q: int(q.get("_order", 10**9)))
        qs, removed_dupes = v3_dedupe_questions(qs)

    debug = {
        "parser": parser_mode,
        "items_extracted": len(desc_items),
        "matching": sum(1 for q in qs if q.get("kind") == "matching"),
        "mcq": sum(1 for q in qs if q.get("kind") == "mcq"),
        "essay": sum(1 for q in qs if q.get("kind") == "essay"),
        "total": len(qs),
        "removed_dupes": removed_dupes,
        "ai_log": ai_log[:60],
    }
    return qs, description_html, debug


def _clean_question_for_response(q: dict) -> dict:
    """Strip internal-only keys (_order, src) before returning to the client."""
    return {
        "question": q.get("question", ""),
        "kind": q.get("kind", "mcq"),
        "options": q.get("options") or [],
        "correct": q.get("correct") or [],
        "multi": bool(q.get("multi", False)),
        "pairs": q.get("pairs") or [],
        "qnum": q.get("qnum"),
    }


# ===========================================================================
# Views
# ===========================================================================

class VerifyCanvasAuthView(APIView):
    """
    POST /api/v1/auth/verify/
    Verify a Canvas token and return the authenticated user's profile.

    Request body:
        { "canvas_base_url": "...", "canvas_token": "..." }

    Response 200:
        { "id": 123, "name": "Jane Doe", "login_id": "jane@example.com" }

    Response 401:
        { "error": "Invalid or expired Canvas token." }
    """

    def post(self, request):
        ser = CanvasAuthSerializer(data=request.data)
        if not ser.is_valid():
            return Response({"error": ser.errors}, status=status.HTTP_400_BAD_REQUEST)

        d = ser.validated_data
        try:
            me = canvas_whoami(d["canvas_base_url"], d["canvas_token"])
        except Exception as exc:
            return Response({"error": str(exc)}, status=status.HTTP_502_BAD_GATEWAY)

        if me is None:
            return Response(
                {"error": "Invalid or expired Canvas token."},
                status=status.HTTP_401_UNAUTHORIZED,
            )
        return Response(me, status=status.HTTP_200_OK)


class CoursesView(APIView):
    """
    GET /api/v1/courses/?canvas_base_url=...&canvas_token=...
    List all courses visible to the supplied Canvas token.

    Query params:
        canvas_base_url  (required)
        canvas_token     (required)

    Response 200:
        [{ "id": 1, "name": "Course Name", "course_code": "CS101" }, ...]
    """

    def get(self, request):
        canvas_base_url = request.query_params.get("canvas_base_url", "").strip()
        canvas_token = request.query_params.get("canvas_token", "").strip()

        if not canvas_base_url or not canvas_token:
            return Response(
                {"error": "canvas_base_url and canvas_token query params are required."},
                status=status.HTTP_400_BAD_REQUEST,
            )
        try:
            courses = list_courses(canvas_base_url, canvas_token)
        except Exception as exc:
            return Response({"error": str(exc)}, status=status.HTTP_502_BAD_GATEWAY)

        simplified = [
            {
                "id": c.get("id"),
                "name": (c.get("name") or c.get("course_code") or f"Course {c.get('id')}").strip(),
                "course_code": c.get("course_code", ""),
            }
            for c in courses
        ]
        return Response(simplified, status=status.HTTP_200_OK)


class ParseDocxView(APIView):
    """
    POST /api/v1/parse/
    Upload a DOCX file and parse it into structured quiz questions.

    Request: multipart/form-data
        file          DOCX file (required)
        parser_mode   "v1" | "v2" | "v3"   (default: "v1")
        openai_api_key   (required when parser_mode == "v3")
        openai_model     (default: "gpt-4.1-mini")
        openai_base_url  (default: "https://api.openai.com")

    Response 200:
        {
          "questions": [ { question, kind, options, correct, multi, pairs, qnum }, ... ],
          "description_html": "<p>...</p>",
          "filename": "my_quiz.docx",
          "parser_mode": "v1",
          "debug": { ... }
        }
    """

    parser_classes = [MultiPartParser, JSONParser]

    def post(self, request):
        ser = ParseRequestSerializer(data=request.data)
        if not ser.is_valid():
            return Response({"error": ser.errors}, status=status.HTTP_400_BAD_REQUEST)

        d = ser.validated_data
        parser_mode = d["parser_mode"]
        uploaded_file = d["file"]
        filename = getattr(uploaded_file, "name", "upload.docx")

        # Build OpenAI config only when needed
        openai_cfg = None
        if parser_mode == "v3":
            api_key = (d.get("openai_api_key") or "").strip()
            if not api_key:
                return Response(
                    {"error": "openai_api_key is required for v3 parser."},
                    status=status.HTTP_400_BAD_REQUEST,
                )
            openai_cfg = OpenAIConfig(
                api_key=api_key,
                model=(d.get("openai_model") or "gpt-4.1-mini").strip(),
                base_url=(d.get("openai_base_url") or "https://api.openai.com").strip(),
            )

        # Write uploaded file to a temp path
        suffix = os.path.splitext(filename)[1] or ".docx"
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                for chunk in uploaded_file.chunks():
                    tmp.write(chunk)
                docx_path = tmp.name

            qs, description_html, debug = _run_parser(docx_path, parser_mode, openai_cfg)
        except ValueError as exc:
            return Response({"error": str(exc)}, status=status.HTTP_400_BAD_REQUEST)
        except Exception as exc:
            return Response(
                {"error": "Parsing failed", "detail": str(exc)},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )
        finally:
            try:
                os.unlink(docx_path)
            except Exception:
                pass

        cleaned_qs = [_clean_question_for_response(q) for q in qs]
        return Response(
            {
                "questions": cleaned_qs,
                "description_html": description_html,
                "filename": filename,
                "parser_mode": parser_mode,
                "debug": debug,
            },
            status=status.HTTP_200_OK,
        )


class ValidateQuestionsView(APIView):
    """
    POST /api/v1/questions/validate/
    Validate a list of questions before uploading to Canvas.
    Returns a list of problems (empty list = all clear).

    Request body:
        { "questions": [ { question, kind, options, correct, multi, pairs }, ... ] }

    Response 200:
        { "valid": true, "problems": [] }
    or
        { "valid": false, "problems": ["Q1: ...", "Q3: ..."] }
    """

    def post(self, request):
        ser = ValidateQuestionsSerializer(data=request.data)
        if not ser.is_valid():
            return Response({"error": ser.errors}, status=status.HTTP_400_BAD_REQUEST)

        questions = ser.validated_data["questions"]
        problems = validate_before_upload(questions)
        return Response(
            {"valid": len(problems) == 0, "problems": problems},
            status=status.HTTP_200_OK,
        )


class UploadQuizView(APIView):
    """
    POST /api/v1/quiz/upload/
    Create a Canvas quiz, upload all questions, and optionally publish it.

    Request body (application/json):
        {
          "canvas_base_url": "...",
          "canvas_token": "...",
          "course_id": "12345",
          "publish": false,
          "settings": {
            "quiz_title": "My Quiz",
            "description_html": "<p>...</p>",
            "shuffle_answers": true,
            "time_limit": 0,
            "allow_multiple_attempts": false,
            "allowed_attempts": 1,
            "scoring_policy": "keep_highest",
            "one_question_at_a_time": false,
            "show_correct_answers": false,
            "access_code_enabled": false,
            "access_code": "",
            "due_at": "",
            "unlock_at": "",
            "lock_at": ""
          },
          "questions": [ { question, kind, options, correct, multi, pairs }, ... ]
        }

    Response 201:
        {
          "quiz_id": 9876,
          "quiz_title": "My Quiz",
          "course_id": "12345",
          "published": false,
          "question_count": 42
        }
    """

    def post(self, request):
        ser = UploadQuizRequestSerializer(data=request.data)
        if not ser.is_valid():
            return Response({"error": ser.errors}, status=status.HTTP_400_BAD_REQUEST)

        d = ser.validated_data
        canvas_base_url = d["canvas_base_url"]
        canvas_token = d["canvas_token"]
        course_id = str(d["course_id"])
        questions = d["questions"]
        settings = d["settings"]
        do_publish = d["publish"]

        # Validate questions first
        problems = validate_before_upload(questions)
        if problems:
            return Response(
                {"error": "Question validation failed", "problems": problems},
                status=status.HTTP_422_UNPROCESSABLE_ENTITY,
            )

        try:
            existing_titles = get_existing_quiz_titles(canvas_base_url, course_id, canvas_token)
            final_title = generate_unique_title(settings["quiz_title"], existing_titles)

            quiz_id = create_canvas_quiz(
                canvas_base_url=canvas_base_url,
                course_id=course_id,
                canvas_token=canvas_token,
                title=final_title,
                description_html=settings.get("description_html", ""),
                settings=settings,
            )

            for q in questions:
                add_question_to_quiz(canvas_base_url, course_id, canvas_token, quiz_id, q)

            if do_publish:
                publish_quiz(canvas_base_url, course_id, canvas_token, quiz_id)

        except RuntimeError as exc:
            # canvas_api raises RuntimeError for 401/403
            http_status = (
                status.HTTP_401_UNAUTHORIZED
                if "401" in str(exc)
                else status.HTTP_403_FORBIDDEN
                if "403" in str(exc)
                else status.HTTP_502_BAD_GATEWAY
            )
            return Response({"error": str(exc)}, status=http_status)
        except Exception as exc:
            return Response(
                {"error": "Canvas upload failed", "detail": str(exc)},
                status=status.HTTP_502_BAD_GATEWAY,
            )

        return Response(
            {
                "quiz_id": quiz_id,
                "quiz_title": final_title,
                "course_id": course_id,
                "published": do_publish,
                "question_count": len(questions),
            },
            status=status.HTTP_201_CREATED,
        )