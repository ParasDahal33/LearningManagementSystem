"""
core/config.py
Environment-based configuration for the Django API.
All values can be overridden via environment variables or a .env file.
"""
from __future__ import annotations

import os

import pytz

TZ_NAME = "Australia/Sydney"
tz = pytz.timezone(TZ_NAME)


def get_env(key: str, default: str = "") -> str:
    return os.environ.get(key, default)


# Canvas defaults (callers should pass these per-request, but useful for testing)
CANVAS_BASE_URL: str = get_env("CANVAS_BASE_URL", "https://learningvault.test.instructure.com/api/v1")
CANVAS_TOKEN: str = get_env("CANVAS_TOKEN", "")

# OpenAI defaults
OPENAI_API_KEY: str = get_env("OPENAI_API_KEY", "")
OPENAI_MODEL: str = get_env("OPENAI_MODEL", "gpt-4.1-mini")
OPENAI_BASE_URL: str = get_env("OPENAI_BASE_URL", "https://api.openai.com")