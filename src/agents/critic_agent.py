"""
critic_agent.py — Step 9: Critic Agent
Reviews draft_response before it's sent to the user.
Uses DeepSeek R1 Distill (reasoning model) for quality checking.
Approves or returns feedback for retry.
"""

import json
import logging
from langchain_core.messages import HumanMessage, SystemMessage

from src.graph.state import AgentState
from src.llm.router import get_llm

logger = logging.getLogger(__name__)

CRITIC_SYSTEM = """You are a strict quality reviewer for an AI assistant called DocuForge.

Review the draft response before it goes to the user. Check for:
1. Accuracy — no hallucinations or made-up facts
2. Completeness — actually answers the user's question/request
3. Clarity — clear and well-structured
4. Relevance — stays on topic, no unnecessary filler
5. For document tasks: confirms if the document task was completed or explains failure clearly
6. For research tasks: citations are present and properly formatted

Respond ONLY with JSON:
{
  "approved": true | false,
  "score": 0-10,
  "issues": ["issue1", "issue2"],  // empty if approved
  "feedback": "Specific improvement instructions for the next attempt (empty if approved)",
  "revised_response": "Your improved version of the response (only if minor fix needed, else empty string)"
}

Be practical — approve if the response is good enough. Don't nitpick style.
Score 7+ = approve. Score < 7 = reject with feedback.
"""


async def critic_agent_node(state: AgentState) -> dict:
    """
    LangGraph node for critic review.
    If approved: passes draft_response as final_response.
    If rejected: returns feedback for the originating agent to retry.
    """
    draft = state.get("draft_response", "")
    intent = state.get("intent", "chat")
    active_agent = state.get("active_agent", "unknown")
    retry_count = state.get("retry_count", 0)

    if not draft:
        return {
            "final_response": "I encountered an error generating a response.",
            "active_agent": "critic_agent",
        }

    # Skip critic on final retry to avoid infinite loops
    if retry_count >= 3:
        logger.info("[CriticAgent] Review budget exhausted — document approved as final.")
        return {
            "final_response": draft,
            "critic_feedback": "Review budget exhausted — draft approved.",
            "active_agent": "critic_agent",
        }

    # Get original user message for context
    last_human = next(
        (m for m in reversed(state["messages"]) if hasattr(m, "type") and m.type == "human"),
        None,
    )
    user_query = last_human.content if last_human else "unknown"

    llm = get_llm(task="critic")  # DeepSeek R1 Distill

    prompt = f"""User's original request:
{user_query}

Intent classified as: {intent}
Agent that produced this: {active_agent}

Draft response to review:
\"\"\"
{draft}
\"\"\"

Review the draft and return your JSON verdict."""

    messages = [
        SystemMessage(content=CRITIC_SYSTEM),
        HumanMessage(content=prompt),
    ]

    try:
        response = await llm.ainvoke(messages)
        raw = response.content.strip()

        # Strip thinking tags (DeepSeek R1 / Qwen QwQ)
        if "<think>" in raw:
            raw = raw.split("</think>")[-1].strip()

        # Strip markdown code fences
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else raw
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.strip()

        result = json.loads(raw)
        approved = result.get("approved", False)
        score = result.get("score", 0)
        feedback = result.get("feedback", "")
        revised = result.get("revised_response", "")

        logger.info(f"[CriticAgent] approved={approved} score={score}")

        if approved:
            # Use revised if provided (minor inline fix), else original draft
            final = revised if revised else draft
            return {
                "final_response": final,
                "critic_feedback": "",
                "active_agent": "critic_agent",
            }
        else:
            # Reject — send feedback back for retry
            logger.info(f"[CriticAgent] Rejected. Feedback: {feedback}")
            return {
                "final_response": "",   # Empty signals retry needed
                "critic_feedback": feedback,
                "active_agent": "critic_agent",
            }

    except (json.JSONDecodeError, KeyError) as e:
        logger.error(f"[CriticAgent] Parse error: {e} — auto-approving draft")
        return {
            "final_response": draft,
            "critic_feedback": "",
            "active_agent": "critic_agent",
        }
