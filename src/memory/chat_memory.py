"""
chat_memory.py — Retrieve relevant past chat history and inject into context.
"""

import logging

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Availability check — onnxruntime (required by fastembed / FAISSStore) may
# fail to load on some Windows installs due to a DLL initialisation error.
# If it's unavailable we silently disable semantic memory; everything else
# in the app continues to work normally.
# ---------------------------------------------------------------------------
_MEMORY_ENABLED = False
try:
    import onnxruntime  # noqa: F401  — just probe the DLL
    from src.memory.faiss_store import FAISSStore
    _MEMORY_ENABLED = True
except Exception as _mem_exc:
    logger.info(
        "[ChatMemory] Semantic memory disabled — onnxruntime is not available "
        f"on this machine ({type(_mem_exc).__name__}). "
        "Document processing is unaffected."
    )
    FAISSStore = None  # type: ignore[assignment,misc]

_stores: dict = {}  # user_id → FAISSStore (singleton per user)


def get_store(user_id: str):
    """Return the FAISSStore for *user_id*, or None if memory is disabled."""
    if not _MEMORY_ENABLED:
        return None
    if user_id not in _stores:
        _stores[user_id] = FAISSStore(user_id)  # type: ignore[misc]
    return _stores[user_id]


async def retrieve_relevant_history(user_id: str, query: str, top_k: int = 5) -> list[dict]:
    """
    Retrieve past messages semantically relevant to the current query.
    Returns list of {text, session_id, role, timestamp, score}.
    Returns an empty list when memory is disabled.
    """
    store = get_store(user_id)
    if store is None:
        return []
    results = store.search(query, top_k=top_k)
    logger.info(f"[ChatMemory] Retrieved {len(results)} memories for user {user_id}")
    return results


async def save_message(user_id: str, session_id: str, role: str, content: str, timestamp: str = ""):
    """Persist a message to FAISS for future retrieval. No-op when memory is disabled."""
    store = get_store(user_id)
    if store is None:
        return
    metadata = {
        "session_id": session_id,
        "role": role,
        "timestamp": timestamp,
    }
    store.add(content, metadata)


def format_history_for_context(memories: list[dict]) -> str:
    """Format retrieved memories as a readable context block for the LLM."""
    if not memories:
        return ""
    lines = ["Relevant past context:"]
    for m in memories:
        role = m.get("role", "unknown").capitalize()
        text = m.get("text", "")
        ts = m.get("timestamp", "")
        date_hint = f" ({ts})" if ts else ""
        lines.append(f"[{role}{date_hint}]: {text}")
    return "\n".join(lines)
