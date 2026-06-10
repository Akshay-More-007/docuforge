"""
End-to-end test: uniformity request on the messy SRM v4 doc through the
document agent, including one critic-style retry with carried-over state
(mirrors the graph's retry loop).
"""
import asyncio
import logging
import sys

from dotenv import load_dotenv

load_dotenv()
sys.path.insert(0, ".")

logging.basicConfig(level=logging.INFO, format="%(name)s - %(message)s")

from src.agents.document_agent import document_agent_node
from src.graph.state import AgentState

SOURCE = r"C:\Users\Batman\Downloads\SRM_Process_Document_v4.docx"

REQ = (
    "Make this SOP document fully uniform and professional: use Arial font "
    "everywhere, make every table header the same red color C00000, keep heading "
    "colors consistent, fix any spelling mistakes, and refresh the RACI matrix "
    "and the process flow chart to match the document content."
)


async def main():
    state = AgentState(
        messages=[],
        intent="doc_task_docx",
        source_doc_path=SOURCE,
        template_doc_path="",
        requirements=[REQ],
        session_id="uniformity_e2e",
        retry_count=0,
        critic_feedback="",
        doc_enhancements={},
    )

    print("\n──────── PASS 1 ────────")
    r1 = await document_agent_node(state)
    print(f"requirements_met: {r1.get('requirements_met')}")
    fixes1 = (r1.get("doc_enhancements") or {}).get("spelling_fixes") or {}
    print(f"fixes carried: {list(fixes1.items())}")

    if r1.get("requirements_met"):
        print("Converged in one pass.")
        return

    print("\n──────── PASS 2 (retry with feedback) ────────")
    state.update(
        retry_count=r1["retry_count"],
        critic_feedback=r1["critic_feedback"],
        doc_enhancements=r1.get("doc_enhancements") or {},
    )
    r2 = await document_agent_node(state)
    print(f"requirements_met: {r2.get('requirements_met')}")
    fixes2 = (r2.get("doc_enhancements") or {}).get("spelling_fixes") or {}
    print(f"fixes carried (should include pass-1 fixes): {list(fixes2.items())}")
    print(f"output: {r2.get('output_doc_path')}")
    print(f"draft_response: {r2.get('draft_response', '')[:300]}")


asyncio.run(main())
