from typing import TypedDict, Annotated
from langgraph.graph import add_messages


class AgentState(TypedDict):
    # Core
    messages: Annotated[list, add_messages]
    user_id: str
    session_id: str

    # Routing
    intent: str          # doc_task_docx | doc_task_pptx | doc_task_pdf | research | chat | code | unclear
    active_agent: str

    # Document task
    source_doc_path: str
    template_doc_path: str
    extracted_content: dict
    extracted_styles: dict
    requirements: list[str]
    requirements_met: bool
    output_doc_path: str
    retry_count: int

    # Research
    search_results: list
    citations: list

    # Memory
    relevant_history: list

    # Quality
    draft_response: str
    critic_feedback: str
    final_response: str
