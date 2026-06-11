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

    The "document" task is routed to Gemini by default: it sends one very large
    prompt (the full block-level document view) that free-tier Groq TPM caps
    reject (413), and granular edit plans need a stronger model than the small
    Groq ones. Set DOC_PROVIDER=groq to force Groq for documents.
    """
    model = TASK_MODEL_MAP.get(task, GROQ_FAST)
    groq_key = os.getenv("GROQ_API_KEY")

    if (
        task == "document"
        and os.getenv("GOOGLE_API_KEY")
        and os.getenv("DOC_PROVIDER", "google").lower() != "groq"
    ):
        logger.debug(f"[Router] task={task} → Gemini chain @ {temperature}")
        return _gemini_chain(temperature, primary_first=True)

    if task == "learn" and os.getenv("GOOGLE_API_KEY"):
        # Learn-mode generation is simpler than document edit plans — start on
        # flash-lite to save the small 2.5-flash daily quota for documents.
        logger.debug(f"[Router] task={task} → Gemini chain (lite first)")
        return _gemini_chain(temperature, primary_first=False)

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


def _get_gemini(temperature: float = 0.1, model: str | None = None) -> ChatGoogleGenerativeAI:
    """Return a cached Gemini client for the given temperature and model."""
    google_key = os.getenv("GOOGLE_API_KEY")
    if not google_key:
        raise EnvironmentError(
            "Neither GROQ_API_KEY nor GOOGLE_API_KEY is set. "
            "At least one LLM provider must be configured in .env"
        )
    model = model or GEMINI_MODEL
    cache_key = (model, temperature)
    if cache_key not in _gemini_clients:
        _gemini_clients[cache_key] = ChatGoogleGenerativeAI(
            model=model,
            temperature=temperature,
            google_api_key=google_key,
        )
    return _gemini_clients[cache_key]


# Free-tier Gemini quotas are PER MODEL (e.g. only 20 req/day on 2.5-flash),
# so sibling models give independent daily budgets. The chain ends on Groq
# Scout — its 30K TPM is the only free Groq limit that fits the large
# document prompts (gpt-oss-120b is capped at 8K TPM → 413).
GEMINI_FALLBACK_MODELS = ["gemini-2.5-flash-lite", "gemini-2.0-flash"]

# Cache: (temperature, primary_first) → RunnableWithFallbacks
_gemini_chains: dict = {}


def _gemini_chain(temperature: float, primary_first: bool = True):
    """Gemini with per-model-quota fallbacks, ending on Groq when available."""
    cache_key = (temperature, primary_first)
    if cache_key in _gemini_chains:
        return _gemini_chains[cache_key]

    if primary_first:
        order = [GEMINI_MODEL] + GEMINI_FALLBACK_MODELS
    else:  # cheap-first for lighter tasks
        order = GEMINI_FALLBACK_MODELS[:1] + [GEMINI_MODEL] + GEMINI_FALLBACK_MODELS[1:]

    llms = [_get_gemini(temperature, m) for m in order]
    if os.getenv("GROQ_API_KEY"):
        llms.append(ChatGroq(
            model=GROQ_LONG,
            temperature=temperature,
            api_key=os.getenv("GROQ_API_KEY"),
            max_retries=2,
        ))

    chain = llms[0].with_fallbacks(llms[1:]) if len(llms) > 1 else llms[0]
    _gemini_chains[cache_key] = chain
    return chain
