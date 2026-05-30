"""
nodes.py — Each agent wrapped as a LangGraph node.
All nodes receive AgentState and return partial state updates.
"""

from src.agents.memory_agent import memory_retrieve_node, memory_save_node
from src.agents.intent_router import intent_router_node
from src.agents.document_agent import document_agent_node
from src.agents.research_agent import research_agent_node
from src.agents.critic_agent import critic_agent_node
from src.agents.fallback_agent import fallback_agent_node
from src.graph.state import AgentState
from src.llm.router import get_llm
from langchain_core.messages import HumanMessage, SystemMessage, AIMessage
import logging

logger = logging.getLogger(__name__)


async def chat_node(state: AgentState) -> dict:
    """
    Direct LLM response for 'chat' and 'code' intents.
    Injects relevant memory context if available.
    """
    llm = get_llm(task="chat")

    system_parts = ["You are DocuForge, a helpful AI assistant specializing in document reformatting and research."]

    # Inject memory context
    history = state.get("relevant_history", [])
    if history:
        from src.memory.chat_memory import format_history_for_context
        context = format_history_for_context(history)
        if context:
            system_parts.append(context)

    messages = [SystemMessage(content="\n\n".join(system_parts))] + list(state["messages"])

    try:
        response = await llm.ainvoke(messages)
        return {
            "draft_response": response.content,
            "active_agent": "chat",
        }
    except Exception as e:
        logger.error(f"[ChatNode] Error: {e}")
        return {
            "draft_response": f"Error generating response: {e}",
            "active_agent": "chat",
        }


# Re-export all nodes for graph_builder
__all__ = [
    "memory_retrieve_node",
    "memory_save_node",
    "intent_router_node",
    "document_agent_node",
    "research_agent_node",
    "critic_agent_node",
    "fallback_agent_node",
    "chat_node",
]
