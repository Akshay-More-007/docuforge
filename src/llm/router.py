"""
router.py — Smart model routing.

Task → Model:
  routing/chat/research/code  → Llama 3.3 70B (Groq, fast)
  document                    → Llama 4 Scout (Groq, 512K context)
  critic/validation           → Qwen QwQ 32B (Groq, reasoning)
  [any, Groq unavailable]     → Gemini 2.5 Flash (Google AI Studio)

Model names can be overridden via environment variables.
"""

import logging
import os
from langchain_groq import ChatGroq
from langchain_google_genai import ChatGoogleGenerativeAI

logger = logging.getLogger(__name__)

# Model identifiers — override via .env if needed
GROQ_FAST   = os.getenv("GROQ_MODEL_FAST",   "llama-3.3-70b-versatile")
GROQ_LONG   = os.getenv("GROQ_MODEL_LONG",   "llama-4-scout-17b-16e-instruct")
GROQ_REASON = os.getenv("GROQ_MODEL_REASON", "qwen-qwq-32b")
GEMINI_MODEL = os.getenv("GEMINI_MODEL",     "gemini-2.5-flash")

TASK_MODEL_MAP: dict[str, str] = {
    "routing":    GROQ_FAST,
    "chat":       GROQ_FAST,
    "research":   GROQ_FAST,
    "code":       GROQ_FAST,
    "document":   GROQ_LONG,
    "critic":     GROQ_REASON,
    "validation": GROQ_REASON,
}

# Cache: (model_name, temperature) → ChatGroq instance
_groq_clients: dict[tuple, ChatGroq] = {}
# Cache: (model_name, temperature) → ChatGoogleGenerativeAI instance
_gemini_clients: dict[tuple, ChatGoogleGenerativeAI] = {}


def get_llm(task: str = "chat", temperature: float = 0.1):
    """
    Returns the appropriate LLM for a given task.
    Falls back to Gemini 2.5 Flash if GROQ_API_KEY is missing.
    Raises EnvironmentError if neither key is set.
    """
    model = TASK_MODEL_MAP.get(task, GROQ_FAST)
    groq_key = os.getenv("GROQ_API_KEY")

    if groq_key:
        cache_key = (model, temperature)
        if cache_key not in _groq_clients:
            _groq_clients[cache_key] = ChatGroq(
                model=model,
                temperature=temperature,
                api_key=groq_key,
                max_retries=2,
            )
        logger.debug(f"[Router] task={task} → {model} @ {temperature} (Groq)")
        return _groq_clients[cache_key]

    # Fallback to Gemini
    logger.info(f"[Router] GROQ_API_KEY not set — falling back to Gemini for task={task}")
    return _get_gemini(temperature)


def _get_gemini(temperature: float = 0.1) -> ChatGoogleGenerativeAI:
    """Return a cached Gemini client for the given temperature."""
    google_key = os.getenv("GOOGLE_API_KEY")
    if not google_key:
        raise EnvironmentError(
            "Neither GROQ_API_KEY nor GOOGLE_API_KEY is set. "
            "At least one LLM provider must be configured in .env"
        )
    cache_key = (GEMINI_MODEL, temperature)
    if cache_key not in _gemini_clients:
        _gemini_clients[cache_key] = ChatGoogleGenerativeAI(
            model=GEMINI_MODEL,
            temperature=temperature,
            google_api_key=google_key,
        )
    return _gemini_clients[cache_key]
