"""
test_graph.py — Step 14: Test full agent flow end-to-end.
Tests: chat, research, and unclear intent flows through the full graph.
"""

from dotenv import load_dotenv
load_dotenv()

import asyncio
import uuid
from langchain_core.messages import HumanMessage
from src.graph.graph_builder import compiled_graph
from src.graph.state import AgentState


def make_state(message: str, **kwargs) -> AgentState:
    return AgentState(
        messages=[HumanMessage(content=message)],
        user_id="testuser",
        session_id=str(uuid.uuid4()),
        intent="",
        active_agent="",
        source_doc_path=kwargs.get("source_doc_path", ""),
        template_doc_path=kwargs.get("template_doc_path", ""),
        extracted_content={},
        extracted_styles={},
        requirements=kwargs.get("requirements", []),
        requirements_met=False,
        output_doc_path="",
        retry_count=0,
        search_results=[],
        citations=[],
        relevant_history=[],
        draft_response="",
        critic_feedback="",
        final_response="",
    )


async def test_chat():
    print("\n" + "="*50)
    print("TEST 1: Chat intent")
    print("="*50)
    state = make_state("What is the difference between supervised and unsupervised learning?")
    result = await compiled_graph.ainvoke(state)
    print(f"Intent:   {result['intent']}")
    print(f"Agent:    {result['active_agent']}")
    print(f"Response: {result['final_response'][:200]}...")
    assert result["intent"] == "chat", f"Expected 'chat', got '{result['intent']}'"
    assert result["final_response"], "final_response is empty"
    print("✅ PASSED")


async def test_research():
    print("\n" + "="*50)
    print("TEST 2: Research intent")
    print("="*50)
    state = make_state("What are the latest AI models released in 2025?")
    result = await compiled_graph.ainvoke(state)
    print(f"Intent:   {result['intent']}")
    print(f"Agent:    {result['active_agent']}")
    print(f"Citations: {len(result.get('citations', []))}")
    print(f"Response: {result['final_response'][:200]}...")
    assert result["intent"] == "research", f"Expected 'research', got '{result['intent']}'"
    assert result["final_response"], "final_response is empty"
    print("✅ PASSED")


async def test_unclear():
    print("\n" + "="*50)
    print("TEST 3: Unclear intent")
    print("="*50)
    state = make_state("uh I need help with something")
    result = await compiled_graph.ainvoke(state)
    print(f"Intent:   {result['intent']}")
    print(f"Agent:    {result['active_agent']}")
    print(f"Response: {result['final_response'][:200]}...")
    assert result["intent"] == "unclear", f"Expected 'unclear', got '{result['intent']}'"
    assert result["final_response"], "final_response is empty"
    print("✅ PASSED")


async def main():
    print("DocuForge — Phase 3 Graph Tests")
    await test_chat()
    await test_research()
    await test_unclear()
    print("\n" + "="*50)
    print("All tests passed. Phase 3 complete.")
    print("="*50)


asyncio.run(main())
