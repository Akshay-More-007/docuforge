"""
memory_agent.py — Step 10: Memory Agent
Runs FIRST in every graph execution.
Retrieves relevant past context and injects into state.
Also saves new messages after the response is final.
"""

import logging
from langchain_core.messages import HumanMessage

from src.graph.state import AgentState
from src.memory.chat_memory import retrieve_relevant_history, save_message, format_history_for_context

logger = logging.getLogger(__name__)


async def memory_retrieve_node(state: AgentState) -> dict:
    """
    Pre-processing node: retrieve relevant past context.
    Runs before Intent Router so context is available to all agents.
    """
    user_id = state.get("user_id", "")
    if not user_id:
        return {"relevant_history": [], "active_agent": "memory_agent"}

    # Get current query
    last_human = next(
        (m for m in reversed(state["messages"]) if isinstance(m, HumanMessage)),
        None,
    )
    if not last_human:
        return {"relevant_history": [], "active_agent": "memory_agent"}

    query = last_human.content if isinstance(last_human.content, str) else str(last_human.content)

    try:
        memories = await retrieve_relevant_history(user_id, query, top_k=5)
        logger.info(f"[MemoryAgent] Retrieved {len(memories)} memories for user {user_id}")
        return {
            "relevant_history": memories,
            "active_agent": "memory_agent",
        }
    except Exception as e:
        logger.warning(f"[MemoryAgent] Retrieval error: {e}")
        return {"relevant_history": [], "active_agent": "memory_agent"}


async def memory_save_node(state: AgentState) -> dict:
    """
    Post-processing node: save the final exchange to FAISS.
    Runs after critic approves (or after max retries exhausted).

    If final_response is empty (critic rejected after budget exhausted),
    fall back to draft_response so the user always gets something.
    """
    user_id = state.get("user_id", "")
    session_id = state.get("session_id", "")

    # Fallback: use draft_response if critic never approved
    final_response = state.get("final_response", "") or state.get("draft_response", "")

    if not user_id or not final_response:
        return {"final_response": final_response}  # at minimum surface the response

    # Ensure final_response is written back into state so chat.py can read it
    state_update: dict = {"final_response": final_response}

    # Get last human message
    last_human = next(
        (m for m in reversed(state["messages"]) if isinstance(m, HumanMessage)),
        None,
    )

    try:
        if last_human:
            user_text = last_human.content if isinstance(last_human.content, str) else str(last_human.content)
            await save_message(user_id, session_id, "user", user_text)

        await save_message(user_id, session_id, "assistant", final_response)
        logger.info(f"[MemoryAgent] Saved exchange to FAISS for user {user_id}")
    except Exception as e:
        logger.debug(f"[MemoryAgent] Save skipped: {e}")

    return state_update
