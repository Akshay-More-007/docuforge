from dotenv import load_dotenv
load_dotenv()

import logging
from langchain_core.messages import HumanMessage, SystemMessage

from src.agents.intent_router import intent_router_node
from src.graph.state import AgentState
from src.llm.router import get_llm

logger = logging.getLogger(__name__)

FALLBACK_SYSTEM = """You are DocuForge, an AI assistant specializing in document reformatting and research.

The user's request was unclear. Ask ONE short, specific clarifying question to understand what they need.
Options to clarify between:
- Reformat a Word document (.docx)
- Create a PowerPoint presentation (.pptx)
- Research / web search
- General chat / coding help

Be friendly and brief. One question only."""


async def fallback_agent_node(state: AgentState) -> dict:
    llm = get_llm(task="chat")

    last_human = next(
        (m for m in reversed(state["messages"]) if isinstance(m, HumanMessage)),
        None,
    )
    user_text = last_human.content if last_human else "unclear request"

    messages = [
        SystemMessage(content=FALLBACK_SYSTEM),
        HumanMessage(content=f"User said: {user_text}\n\nAsk a clarifying question."),
    ]

    try:
        response = await llm.ainvoke(messages)
        return {"draft_response": response.content, "active_agent": "fallback_agent"}
    except Exception as e:
        return {
            "draft_response": "Could you clarify — document reformatting, research, or something else?",
            "active_agent": "fallback_agent",
        }