"""
edges.py — Conditional routing logic.

Design:
- After intent: route to the right specialist agent.
- After document: always go to critic (no mid-pipeline loops — retry budget tracked in retry_count).
- After critic:
    - approved (final_response non-empty) → memory_save → END
    - rejected (final_response empty) AND retry budget remains → retry the originating agent
    - rejected AND retry budget exhausted → memory_save (pass draft_response as final)
"""

from src.graph.state import AgentState
import logging

logger = logging.getLogger(__name__)

MAX_RETRIES = 3


def route_after_intent(state: AgentState) -> str:
    intent = state.get("intent", "unclear")
    routing = {
        "doc_task_docx": "document_agent",
        "doc_task_pptx": "document_agent",
        "doc_task_pdf":  "document_agent",   # PDF source → docx/pptx output
        "research":      "research_agent",
        "chat":          "chat",
        "code":          "chat",
        "unclear":       "fallback_agent",
    }
    next_node = routing.get(intent, "fallback_agent")
    logger.info(f"[Edge] intent={intent} → {next_node}")
    return next_node


def route_after_document(state: AgentState) -> str:
    """After document agent runs, always go to critic for quality review."""
    return "critic_agent"


def route_after_critic(state: AgentState) -> str:
    """
    - If final_response is non-empty → approved → save to memory → END.
    - If final_response is empty (critic rejected):
        - Retry budget remaining → send back to originating agent.
        - Budget exhausted → force to memory_save (draft_response used as fallback).
    """
    final = state.get("final_response", "")
    retry_count = state.get("retry_count", 0)

    if final:
        # Approved — write to memory and finish
        logger.info("[Edge] Critic approved → memory_save")
        return "memory_save"

    # Critic rejected
    if retry_count >= MAX_RETRIES:
        logger.warning(f"[Edge] Critic rejected but retry_count={retry_count} — forcing memory_save")
        return "memory_save"

    # Route back to whichever agent produced the draft
    active = state.get("active_agent", "")
    intent  = state.get("intent", "chat")

    if active == "document_agent" or intent in ("doc_task_docx", "doc_task_pptx", "doc_task_pdf"):
        logger.info(f"[Edge] Critic rejected (retry {retry_count}) → document_agent")
        return "document_agent"
    elif active == "research_agent" or intent == "research":
        logger.info(f"[Edge] Critic rejected (retry {retry_count}) → research_agent")
        return "research_agent"
    elif active == "fallback_agent" or intent == "unclear":
        logger.info(f"[Edge] Critic rejected (retry {retry_count}) → fallback_agent")
        return "fallback_agent"
    else:
        logger.info(f"[Edge] Critic rejected (retry {retry_count}) → chat")
        return "chat"
