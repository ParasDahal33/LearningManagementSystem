"""
parsers/mcq_parser.py
Multiple-choice and essay question parsing for all parser versions (v1, v2c, v3).
"""

from __future__ import annotations

from .mcq_v1 import (
    NOISE_RE_V1,
    QUESTION_CMD_INNER_RE_V1,
    COMMAND_QUESTION_RE_V1,
    RUBRIC_START_RE_V1,
    ESSAY_GUIDE_RE_V1,
    parse_mcq_questions_v1,
    parse_essay_questions_v1,
)

from .mcq_v2 import (
    V2C_NOISE_RE,
    V2C_QUESTION_CMD_INNER_RE,
    V2C_COMMAND_QUESTION_RE,
    V2C_RUBRIC_START_RE,
    V2C_ESSAY_GUIDE_RE,
    V2C_MATCHING_STEM_RE,
    v2c_merge_dangling_question_lines,
    v2c_parse_mcq_questions,
    v2c_parse_essay_questions,
)

from .mcq_v3 import (
    V3_IGNORE_LINE_RE,
    V3_IGNORE_SECTION_RE,
    V3_IGNORE_TABLE_RE,
    V3_COOKERY_METHOD_WORD_RE,
    V3_QUESTION_START_RE,
    V3_OPTION_LINE_RE,
    _v3_looks_like_question_start,
    v3_parse_essay_questions_rule_based,
    v3_filter_items_for_ai,
)