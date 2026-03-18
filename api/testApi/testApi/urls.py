"""
canvas_quiz_api/urls.py
Root URL configuration — mounts the quiz app under /api/v1/.
"""
from django.urls import path, include
from .views import (
    VerifyCanvasAuthView,
    CoursesView,
    ParseDocxView,
    ValidateQuestionsView,
    UploadQuizView,
)

urlpatterns = [
    path("api/auth/verify/", VerifyCanvasAuthView.as_view(), name="auth-verify"),
    path("api/courses/", CoursesView.as_view(), name="courses-list"),
    path("api/parse/", ParseDocxView.as_view(), name="parse-docx"),
    path("api/questions/validate/", ValidateQuestionsView.as_view(), name="questions-validate"),
    path("api/quiz/upload/", UploadQuizView.as_view(), name="quiz-upload"),
]