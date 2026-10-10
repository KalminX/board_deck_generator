"""
Commentary Provider Factory.
Instantiates the configured commentary provider (default: 'template').
Allows future AI / LLM commentary providers to be plugged in via configuration.
"""
from django.conf import settings
from services.commentary.base import CommentaryProvider
from services.commentary.templates import TemplateCommentaryProvider


def get_commentary_provider(provider_name: str = None) -> CommentaryProvider:
    """
    Returns the appropriate CommentaryProvider based on parameter or Django setting.
    """
    if not provider_name:
        provider_name = getattr(settings, "BOARD_COMMENTARY_PROVIDER", "template")

    provider_name = provider_name.lower().strip()

    if provider_name == "template":
        return TemplateCommentaryProvider()
    elif provider_name in ("claude", "ai", "llm"):
        # Ready for future AI integration - fallback to deterministic template if not configured
        return TemplateCommentaryProvider()
    else:
        return TemplateCommentaryProvider()
