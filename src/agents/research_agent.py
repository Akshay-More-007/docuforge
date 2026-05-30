"""
research_agent.py — Step 8: Research Agent
Tavily web search + LLM synthesis + structured citations.
"""

import logging
import os
from tavily import AsyncTavilyClient
from langchain_core.messages import HumanMessage, SystemMessage

from src.graph.state import AgentState
from src.llm.router import get_llm

logger = logging.getLogger(__name__)

RESEARCH_SYSTEM = """You are a research assistant with access to live web search results.

Synthesize the search results into a clear, accurate answer.
- Be concise but complete
- Cite sources inline using [1], [2], etc. format
- Note the date of information when relevant
- If results conflict, say so
- Never fabricate information not in the search results
"""


async def research_agent_node(state: AgentState) -> dict:
    """
    LangGraph node for research tasks.
    1. Extract search query from user message
    2. Search via Tavily
    3. Synthesize with LLM + citations
    """
    # Get latest user message
    last_human = next(
        (m for m in reversed(state["messages"]) if hasattr(m, "type") and m.type == "human"),
        None,
    )
    if not last_human:
        return {"draft_response": "No query found.", "active_agent": "research_agent"}

    query = last_human.content if isinstance(last_human.content, str) else str(last_human.content)

    # ── Step 1: Tavily search ─────────────────────────────────────────────────
    api_key = os.getenv("TAVILY_API_KEY")
    if not api_key:
        return {
            "draft_response": "Tavily API key not configured. Set TAVILY_API_KEY in .env",
            "active_agent": "research_agent",
            "search_results": [],
            "citations": [],
        }

    try:
        client = AsyncTavilyClient(api_key=api_key)
        search_response = await client.search(
            query=query,
            search_depth="advanced",
            max_results=5,
            include_answer=True,
        )
    except Exception as e:
        logger.error(f"[ResearchAgent] Tavily error: {e}")
        return {
            "draft_response": f"Search failed: {e}",
            "active_agent": "research_agent",
            "search_results": [],
            "citations": [],
        }

    results = search_response.get("results", [])
    quick_answer = search_response.get("answer", "")

    if not results:
        return {
            "draft_response": "No search results found for that query.",
            "active_agent": "research_agent",
            "search_results": [],
            "citations": [],
        }

    # ── Step 2: Format results for LLM ───────────────────────────────────────
    formatted_results = []
    citations = []
    for i, r in enumerate(results, 1):
        title = r.get("title", "")
        url = r.get("url", "")
        content = r.get("content", "")[:500]  # Cap per source
        published = r.get("published_date", "")

        formatted_results.append(
            f"[{i}] {title}\nURL: {url}\nDate: {published}\n{content}"
        )
        citations.append({
            "index": i,
            "title": title,
            "url": url,
            "published_date": published,
        })

    results_text = "\n\n---\n\n".join(formatted_results)

    # ── Step 3: LLM synthesis ─────────────────────────────────────────────────
    llm = get_llm(task="research")  # Llama 3.3 70B

    prompt = f"""User query: {query}

Tavily quick answer: {quick_answer}

Search results:
{results_text}

Synthesize these results into a comprehensive answer with inline citations [1], [2], etc."""

    messages = [
        SystemMessage(content=RESEARCH_SYSTEM),
        HumanMessage(content=prompt),
    ]

    try:
        response = await llm.ainvoke(messages)
        draft = response.content

        # Append source list
        source_list = "\n\n**Sources:**\n" + "\n".join(
            f"[{c['index']}] [{c['title']}]({c['url']})" + (f" — {c['published_date']}" if c['published_date'] else "")
            for c in citations
        )
        draft += source_list

        logger.info(f"[ResearchAgent] Synthesized {len(results)} results for: {query[:60]}")

        return {
            "search_results": results,
            "citations": citations,
            "draft_response": draft,
            "active_agent": "research_agent",
        }

    except Exception as e:
        logger.error(f"[ResearchAgent] LLM synthesis error: {e}")
        return {
            "search_results": results,
            "citations": citations,
            "draft_response": f"Search completed but synthesis failed: {e}",
            "active_agent": "research_agent",
        }
