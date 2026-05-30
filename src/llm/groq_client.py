# DEPRECATED — this file is NOT used by DocuForge.
# All LLM access goes through src/llm/router.py which caches clients
# and handles Groq → Gemini fallback automatically.
#
# Model names here were also inconsistent with router.py.
# Kept only to avoid import errors if anything references it accidentally.

from src.llm.router import get_llm  # noqa: F401 — re-export for any stale imports
