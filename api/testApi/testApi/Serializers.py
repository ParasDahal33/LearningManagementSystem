"""
quiz/serializers.py
DRF serializers — validate incoming request data and shape outgoing responses.
"""
from rest_framework import serializers


# ---------------------------------------------------------------------------
# Canvas auth
# ---------------------------------------------------------------------------
class CanvasAuthSerializer(serializers.Serializer):
    canvas_base_url = serializers.URLField()
    canvas_token = serializers.CharField()


# ---------------------------------------------------------------------------
# Parse request
# ---------------------------------------------------------------------------
class ParseRequestSerializer(serializers.Serializer):
    PARSER_CHOICES = [
        ("v1", "v1 (rule-based)"),
        ("v2", "v2 (rule-based)"),
        ("v3", "v3 (AI-hybrid)"),
    ]
    file = serializers.FileField(help_text="DOCX file to parse")
    parser_mode = serializers.ChoiceField(choices=PARSER_CHOICES, default="v1")
    openai_api_key = serializers.CharField(required=False, allow_blank=True, default="")
    openai_model = serializers.CharField(required=False, allow_blank=True, default="gpt-4.1-mini")
    openai_base_url = serializers.URLField(required=False, default="https://api.openai.com")


# ---------------------------------------------------------------------------
# Question shapes (used inside other serializers)
# ---------------------------------------------------------------------------
class PairSerializer(serializers.Serializer):
    left = serializers.CharField()
    right = serializers.CharField()


class QuestionSerializer(serializers.Serializer):
    KIND_CHOICES = [("mcq", "MCQ"), ("essay", "Essay"), ("matching", "Matching")]
    question = serializers.CharField()
    kind = serializers.ChoiceField(choices=KIND_CHOICES)
    options = serializers.ListField(child=serializers.CharField(), required=False, default=list)
    correct = serializers.ListField(child=serializers.IntegerField(), required=False, default=list)
    multi = serializers.BooleanField(default=False)
    pairs = serializers.ListField(child=PairSerializer(), required=False, default=list)
    qnum = serializers.CharField(required=False, allow_null=True, default=None)


# ---------------------------------------------------------------------------
# Parse response
# ---------------------------------------------------------------------------
class ParseResponseSerializer(serializers.Serializer):
    questions = QuestionSerializer(many=True)
    description_html = serializers.CharField(allow_blank=True)
    filename = serializers.CharField()
    parser_mode = serializers.CharField()
    debug = serializers.DictField(required=False)


# ---------------------------------------------------------------------------
# Quiz settings
# ---------------------------------------------------------------------------
class QuizSettingsSerializer(serializers.Serializer):
    SCORING_CHOICES = [("keep_highest", "Keep Highest"), ("keep_latest", "Keep Latest")]

    quiz_title = serializers.CharField()
    description_html = serializers.CharField(allow_blank=True, default="")
    shuffle_answers = serializers.BooleanField(default=True)
    one_question_at_a_time = serializers.BooleanField(default=False)
    show_correct_answers = serializers.BooleanField(default=False)
    time_limit = serializers.IntegerField(min_value=0, max_value=1440, default=0)
    allow_multiple_attempts = serializers.BooleanField(default=False)
    allowed_attempts = serializers.IntegerField(min_value=1, max_value=20, default=1)
    scoring_policy = serializers.ChoiceField(choices=SCORING_CHOICES, default="keep_highest")
    access_code_enabled = serializers.BooleanField(default=False)
    access_code = serializers.CharField(allow_blank=True, default="")
    due_at = serializers.CharField(allow_blank=True, default="")
    unlock_at = serializers.CharField(allow_blank=True, default="")
    lock_at = serializers.CharField(allow_blank=True, default="")


# ---------------------------------------------------------------------------
# Upload quiz request
# ---------------------------------------------------------------------------
class UploadQuizRequestSerializer(serializers.Serializer):
    canvas_base_url = serializers.URLField()
    canvas_token = serializers.CharField()
    course_id = serializers.CharField()
    settings = QuizSettingsSerializer()
    questions = QuestionSerializer(many=True)
    publish = serializers.BooleanField(default=False)


# ---------------------------------------------------------------------------
# Upload quiz response
# ---------------------------------------------------------------------------
class UploadQuizResponseSerializer(serializers.Serializer):
    quiz_id = serializers.IntegerField()
    quiz_title = serializers.CharField()
    course_id = serializers.CharField()
    published = serializers.BooleanField()
    question_count = serializers.IntegerField()


# ---------------------------------------------------------------------------
# Validate request
# ---------------------------------------------------------------------------
class ValidateQuestionsSerializer(serializers.Serializer):
    questions = QuestionSerializer(many=True)


# ---------------------------------------------------------------------------
# Course list response item
# ---------------------------------------------------------------------------
class CourseSerializer(serializers.Serializer):
    id = serializers.IntegerField()
    name = serializers.CharField()
    course_code = serializers.CharField(allow_blank=True, allow_null=True)