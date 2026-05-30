"""
graph_builder.py — Build and compile the full DocuForge LangGraph graph.

Flow:
  memory_retrieve → intent_router → [document_agent | research_agent | chat | fallback_agent]
                                           ↓
                                       critic_agent  (retry loop up to 3×)
                                           ↓
                                       memory_save → END
"""
from langgraph.graph import StateGraph, END

from src.graph.state import AgentState
from src.graph.nodes import (
    memory_retrieve_node,
    memory_save_node,
    intent_router_node,
    document_agent_node,
    research_agent_node,
    critic_agent_node,
    fallback_agent_node,
    chat_node,
)
from src.graph.edges import (
    route_after_intent,
    route_after_critic,
    route_after_document,
)


def build_graph():
    """
    Build and compile the DocuForge agent graph.
    Returns a compiled LangGraph runnable.
    """
    graph = StateGraph(AgentState)

    # ── Register nodes ────────────────────────────────────────────────────────
    graph.add_node("memory_retrieve",  memory_retrieve_node)
    graph.add_node("intent_router",    intent_router_node)
    graph.add_node("document_agent",   document_agent_node)
    graph.add_node("research_agent",   research_agent_node)
    graph.add_node("chat",             chat_node)
    graph.add_node("fallback_agent",   fallback_agent_node)
    graph.add_node("critic_agent",     critic_agent_node)
    graph.add_node("memory_save",      memory_save_node)

    # ── Entry point ───────────────────────────────────────────────────────────
    graph.set_entry_point("memory_retrieve")

    # ── Fixed edges ───────────────────────────────────────────────────────────
    graph.add_edge("memory_retrieve", "intent_router")

    # ── After intent router → specialist agent ────────────────────────────────
    graph.add_conditional_edges(
        "intent_router",
        route_after_intent,
        {
            "document_agent": "document_agent",
            "research_agent": "research_agent",
            "chat":           "chat",
            "fallback_agent": "fallback_agent",
        },
    )

    # ── Document → critic (always; retry handled in route_after_critic) ───────
    graph.add_conditional_edges(
        "document_agent",
        route_after_document,
        {
            "critic_agent": "critic_agent",
        },
    )

    # ── Research + chat + fallback → critic ───────────────────────────────────
    graph.add_edge("research_agent", "critic_agent")
    graph.add_edge("chat",           "critic_agent")
    graph.add_edge("fallback_agent", "critic_agent")

    # ── Critic: approve → memory_save | reject → retry agent ─────────────────
    graph.add_conditional_edges(
        "critic_agent",
        route_after_critic,
        {
            "memory_save":    "memory_save",
            "document_agent": "document_agent",
            "research_agent": "research_agent",
            "chat":           "chat",
            "fallback_agent": "fallback_agent",
        },
    )

    # ── End ───────────────────────────────────────────────────────────────────
    graph.add_edge("memory_save", END)

    return graph.compile(checkpointer=None)


# Singleton — import this in chat.py
compiled_graph = build_graph()
