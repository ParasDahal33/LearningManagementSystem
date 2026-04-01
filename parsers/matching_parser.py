"""
parsers/matching_parser.py
Matching / drag-and-drop question parsing facade.
Exports functions for all parser versions (v1, v1-exact, v3).
"""

from __future__ import annotations

from .matching_v1 import (
    MATCHING_STEM_RE,
    parse_matching_questions_doc_order,
    parse_matching_questions_doc_order_v1_exact,
)
from .matching_v3 import (
    V3_MATCHING_STEM_RE,
    v3_parse_matching_questions_doc_order,
    v3_parse_table_defined_terms_as_essays,
    v3_parse_table_characteristics_as_essays,
    v3_collect_ignore_texts_from_forced_tables,
)