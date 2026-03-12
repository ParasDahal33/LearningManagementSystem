"""
core/config.py
Handles secrets loading, timezone configuration, and Streamlit session state initialization.
"""

from __future__ import annotations

import os
import tomllib

import pytz
import streamlit as st

# ---------------------------------------------------------------------------
# Timezone
# ---------------------------------------------------------------------------
TZ_NAME = "Australia/Sydney"
tz = pytz.timezone(TZ_NAME)


# ---------------------------------------------------------------------------
# Secrets
# ---------------------------------------------------------------------------
def safe_load_secrets_toml() -> dict[str, str]:
    """Load secrets from .streamlit/secrets.toml if it exists."""
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        # secrets.toml lives two levels up (project_root/.streamlit/)
        secrets_path = os.path.join(base_dir, "..", ".streamlit", "secrets.toml")
        if not os.path.exists(secrets_path):
            return {}
        with open(secrets_path, "rb") as f:
            data = tomllib.load(f)
        out: dict[str, str] = {}
        for k in [
            "CANVAS_BASE_URL",
            "CANVAS_TOKEN",
            "OPENAI_API_KEY",
            "OPENAI_MODEL",
            "OPENAI_BASE_URL",
        ]:
            v = data.get(k)
            if isinstance(v, str):
                out[k] = v
        return out
    except Exception:
        return {}


_LOCAL_SECRETS: dict[str, str] = safe_load_secrets_toml()


# ---------------------------------------------------------------------------
# Session-state helpers
# ---------------------------------------------------------------------------
def ss_init(key: str, value) -> None:
    """Initialise a Streamlit session-state key only if it does not yet exist."""
    if key not in st.session_state:
        st.session_state[key] = value


def init_session_state() -> None:
    """Call once at app startup to set all default session-state values."""
    ss_init("logged_in", False)
    ss_init("me", None)
    ss_init(
        "canvas_token",
        os.getenv("CANVAS_TOKEN", "") or _LOCAL_SECRETS.get("CANVAS_TOKEN", ""),
    )
    ss_init(
        "canvas_base_url",
        os.getenv("CANVAS_BASE_URL", "")
        or _LOCAL_SECRETS.get(
            "CANVAS_BASE_URL",
            "https://learningvault.test.instructure.com/api/v1",
        ),
    )
    ss_init("courses_cache", None)
    ss_init("selected_course_id", None)

    ss_init("docx_filename", None)
    ss_init("description_html", "")
    ss_init("questions", [])
    ss_init("parsed_ok", False)
    ss_init("parse_run_id", 0)
    ss_init("last_parser_mode", None)
    ss_init("questions_page_size", 10)
    ss_init("questions_page", 1)

    ss_init(
        "details",
        {
            "shuffle_answers": True,
            "time_limit": 0,
            "allow_multiple_attempts": False,
            "allowed_attempts": 2,
            "scoring_policy": "keep_highest",
            "one_question_at_a_time": False,
            "show_correct_answers": False,
            "access_code_enabled": False,
            "access_code": "",
            "due_at": "",
            "unlock_at": "",
            "lock_at": "",
        },
    )

    ss_init(
        "openai_api_key",
        os.getenv("OPENAI_API_KEY", "") or _LOCAL_SECRETS.get("OPENAI_API_KEY", ""),
    )
    ss_init(
        "openai_model",
        os.getenv("OPENAI_MODEL", "") or _LOCAL_SECRETS.get("OPENAI_MODEL", "gpt-4.1-mini"),
    )
    ss_init(
        "openai_base_url",
        os.getenv("OPENAI_BASE_URL", "")
        or _LOCAL_SECRETS.get("OPENAI_BASE_URL", "https://api.openai.com"),
    )