from dotenv import load_dotenv
load_dotenv()

import asyncio
from langchain_core.messages import HumanMessage
from src.graph.state import AgentState
from src.agents.intent_router import intent_router_node

state = AgentState(
    messages=[HumanMessage(content="Reformat this SOP to match the template")],
    user_id="test", session_id="s1", intent="", active_agent="",
    source_doc_path="", template_doc_path="", extracted_content={},
    extracted_styles={}, requirements=[], requirements_met=False,
    output_doc_path="", retry_count=0, search_results=[], citations=[],
    relevant_history=[], draft_response="", critic_feedback="", final_response=""
)

result = asyncio.run(intent_router_node(state))
print(result)