"""
Intent Router — Step 6
Classifies user messages into: doc_task_docx | doc_task_pptx | research | chat | code | unclear
Uses Llama 3.3 70B (fast, sufficient for classification).
"""

import json
import logging
from langchain_core.messages import HumanMessage, SystemMessage
from src.graph.state import AgentState
from src.llm.router import get_llm

logger = logging.getLogger(__name__)

INTENT_SYSTEM_PROMPT = """You are an intent classifier for DocuForge, a document reformatting + research chatbot.

Classify the user's latest message into EXACTLY ONE of these intents:

- doc_task_docx  → User wants to reformat, convert, or process a Word (.docx) document
- doc_task_pptx  → User wants to create or reformat a PowerPoint (.pptx) presentation
- doc_task_pdf   → User has uploaded a PDF and wants to extract, summarise, convert, or reformat it
- research       → User wants to search the web, find recent information, or research a topic
- chat           → General conversation, questions the assistant can answer from knowledge
- code           → User wants code written, debugged, or explained
- unclear        → Ambiguous — needs clarification before proceeding

Rules:
- If a .docx file was uploaded → lean toward doc_task_docx
- If a .pptx file was uploaded → lean toward doc_task_pptx
- If a .pdf file was uploaded → lean toward doc_task_pdf
- "reformat", "fix formatting", "apply template", "clean up", "restructure" → doc_task_docx
- "make slides", "presentation", "PowerPoint", "convert to PPT" → doc_task_pptx
- "convert PDF", "extract from PDF", "summarise PDF" → doc_task_pdf
- "search", "find", "latest", "recent", "news", "current" → research
- Respond ONLY with valid JSON. No preamble.

Response format:
{
  "intent": "<one of the 7 intents>",
  "confidence": <0.0–1.0>,
  "reasoning": "<one sentence>"
}
"""


async def intent_router_node(state: AgentState) -> dict:
    """
    Classifies the latest user message and updates state.intent.
    Uses Llama 3.3 70B via Groq (fast model — routing doesn't need heavy reasoning).
    """
    llm = get_llm(task="routing")

    # Get last user message
    last_human = next(
        (m for m in reversed(state["messages"]) if isinstance(m, HumanMessage)),
        None,
    )
    if not last_human:
        return {"intent": "unclear", "active_agent": "intent_router"}

    user_text = last_human.content if isinstance(last_human.content, str) else str(last_human.content)

    # Include file context hint if paths are set
    context_hints = []
    if state.get("source_doc_path"):
        context_hints.append(f"User has uploaded a source document: {state['source_doc_path']}")
    if state.get("template_doc_path"):
        context_hints.append(f"User has uploaded a template document: {state['template_doc_path']}")

    user_prompt = user_text
    if context_hints:
        user_prompt = "\n".join(context_hints) + "\n\nUser message: " + user_text

    messages = [
        SystemMessage(content=INTENT_SYSTEM_PROMPT),
        HumanMessage(content=user_prompt),
    ]

    try:
        response = await llm.ainvoke(messages)
        raw = response.content.strip()

        # Strip markdown fences if present
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()

        parsed = json.loads(raw)
        intent = parsed.get("intent", "unclear")
        confidence = parsed.get("confidence", 0.0)
        reasoning = parsed.get("reasoning", "")

        logger.info(f"[IntentRouter] intent={intent} confidence={confidence} | {reasoning}")

        # Low confidence → route to unclear
        if confidence < 0.5:
            intent = "unclear"

        return {
            "intent": intent,
            "active_agent": "intent_router",
        }

    except (json.JSONDecodeError, KeyError) as e:
        logger.warning(f"[IntentRouter] Parse error: {e} — defaulting to unclear")
        return {"intent": "unclear", "active_agent": "intent_router"}
