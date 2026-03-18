"""
quiz/exceptions.py
Centralised error responses so every endpoint returns a consistent JSON shape:
  { "error": "...", "detail": "..." }
"""
from rest_framework.views import exception_handler
from rest_framework.response import Response
from rest_framework import status


def custom_exception_handler(exc, context):
    """
    Wrap DRF's default handler so unhandled Python exceptions also get a
    structured JSON body instead of an HTML 500 page.
    """
    response = exception_handler(exc, context)

    if response is not None:
        # DRF already handled it — normalise the shape
        data = response.data
        if isinstance(data, dict) and "error" not in data:
            detail = data.get("detail", str(data))
            response.data = {"error": str(detail)}
        return response

    # Unhandled exception — return 500
    return Response(
        {"error": "Internal server error", "detail": str(exc)},
        status=status.HTTP_500_INTERNAL_SERVER_ERROR,
    )