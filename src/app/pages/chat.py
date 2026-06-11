"""
chat.py — Main chat page.

Key design points:
- Two-phase input handling: submitting a message appends it + sets
  `pending_input`, then reruns so the user's message appears IMMEDIATELY.
  The graph then runs on the next script pass behind a spinner/typing
  indicator (process_pending) instead of silently blocking the first render.
- Graph runs in a dedicated thread with asyncio.run() to stay completely
  isolated from Streamlit's own event loop. Using nest_asyncio caused
  Streamlit to process buffered WebSocket messages mid-execution, which
  triggered spurious reruns and lost session_state.messages.
- final_response falls back to draft_response if critic never set it
"""

import uuid
import asyncio
import concurrent.futures
from pathlib import Path

import streamlit as st

from src.auth.guards import require_auth
from src.app.components.sidebar import render_sidebar
from src.app.components.message_bubble import render_messages, render_thinking
from src.app.compat import rerun
from src.graph.graph_builder import compiled_graph
from src.graph.state import AgentState
from langchain_core.messages import HumanMessage


SUGGESTIONS = [
    ("🧹", "Clean up my document", "Fix all spelling, grammar and casing mistakes, "
     "apply every inline review note, and keep the original theme and formatting."),
    ("📊", "Make tables uniform", "Make this document fully uniform and professional: "
     "consistent fonts, matching table headers, and clean consistent styling."),
    ("🎨", "Polish my deck", "Clean up this deck professionally: fix every spelling and "
     "casing issue and make all table slides follow the same clean uniform design."),
]


def render_chat():
    require_auth()

    if not st.session_state.get("current_session_id"):
        st.session_state.current_session_id = str(uuid.uuid4())
    if not st.session_state.get("messages"):
        st.session_state.messages = []

    render_sidebar()

    st.markdown(
        "<div style='padding: 24px 0 8px 0;'>"
        "<span class='brand' style='font-size:16px;'>⚡ DocuForge</span></div>",
        unsafe_allow_html=True,
    )

    if not st.session_state.messages:
        _render_welcome()

    render_messages(st.session_state.messages)

    # While the graph is running, show the typing indicator under the
    # user's message so the app never looks frozen.
    if st.session_state.get("pending_input"):
        render_thinking()
    else:
        _render_output_card()


def render_input():
    """Called from main.py at top level — must be outside containers."""
    return st.chat_input("Ask anything, or describe what to do with your document...")


def handle_input(user_input: str):
    """Phase 1: record the user message and rerun so it renders instantly."""
    # Defensive init — guards against edge cases where render_chat() was
    # interrupted (e.g. token refresh rerun) before messages was initialised.
    if "messages" not in st.session_state:
        st.session_state.messages = []

    st.session_state.messages.append({
        "role": "user",
        "content": user_input,
        "citations": [],
    })
    st.session_state.pending_input = user_input
    rerun()


def process_pending():
    """Phase 2: run the agent graph for the recorded pending message."""
    user_input = st.session_state.get("pending_input")
    if not user_input:
        return

    initial_state = AgentState(
        messages=[HumanMessage(content=user_input)],
        user_id=st.session_state.get("user_id", ""),
        session_id=st.session_state.get("current_session_id", str(uuid.uuid4())),
        intent="",
        active_agent="",
        source_doc_path=st.session_state.get("source_doc_path") or "",
        template_doc_path=st.session_state.get("template_doc_path") or "",
        extracted_content={},
        extracted_styles={},
        requirements=[user_input],
        requirements_met=False,
        output_doc_path="",
        retry_count=0,
        doc_enhancements={},
        search_results=[],
        citations=[],
        relevant_history=[],
        draft_response="",
        critic_feedback="",
        final_response="",
    )

    has_doc = bool(st.session_state.get("source_doc_path"))
    spinner_text = (
        "Forging your document — extracting, editing, validating…"
        if has_doc else "Thinking…"
    )

    output_path = ""
    try:
        # Run the graph in a dedicated thread with its own event loop so it is
        # fully isolated from Streamlit's event loop.  Using nest_asyncio caused
        # Streamlit to process buffered WebSocket messages during await points,
        # which triggered spurious script reruns and wiped session_state.
        with st.spinner(spinner_text):
            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
                future = pool.submit(
                    asyncio.run,
                    compiled_graph.ainvoke(initial_state, {"recursion_limit": 30}),
                )
                result = future.result()   # blocks Streamlit's script thread cleanly

        # final_response is set by memory_save_node (with draft fallback)
        final = (
            result.get("final_response")
            or result.get("draft_response")
            or "No response generated."
        )
        citations = result.get("citations", [])
        output_path = result.get("output_doc_path", "")
    except Exception as e:
        final = f"⚠️ Error: {e}"
        citations = []

    # Defensive re-init in case a Streamlit rerun cleared state while
    # the graph thread was running (should not happen now, but be safe).
    if "messages" not in st.session_state:
        st.session_state.messages = []

    st.session_state.messages.append({
        "role": "assistant",
        "content": final,
        "citations": citations,
    })

    if output_path:
        st.session_state["last_output_path"] = output_path

    st.session_state.pending_input = None
    rerun()


def _render_output_card():
    """Styled 'output ready' card with a download button, under the chat."""
    output_path = st.session_state.get("last_output_path")
    if not output_path:
        return
    p = Path(output_path)
    if not p.exists():
        return

    size_kb = p.stat().st_size / 1024
    size_str = f"{size_kb / 1024:.1f} MB" if size_kb > 1024 else f"{size_kb:.0f} KB"
    icon = "📊" if p.suffix.lower() == ".pptx" else "📄"

    st.markdown(
        f"""
        <div class='output-card'>
            <div class='output-card-row'>
                <span class='output-card-icon'>{icon}</span>
                <span>
                    <span class='doc-name'>{p.name}</span><br>
                    <span class='output-card-meta'>
                        <span class='badge badge-success'>READY</span>&nbsp; {size_str}
                    </span>
                </span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with open(str(p), "rb") as f:
        st.download_button(
            label=f"⬇️  Download {p.name}",
            data=f.read(),
            file_name=p.name,
            use_container_width=False,
            type="primary",
        )


def _render_welcome():
    st.markdown("""
    <div style="text-align:center; padding:48px 32px 24px 32px; color:#444;">
        <div style="font-family:'Space Mono',monospace; font-size:32px; color:#2a2a2a; margin-bottom:16px;">⚡</div>
        <p style="font-size:13px; color:#555; max-width:480px; margin:0 auto; line-height:1.8;">
            Upload a document in the sidebar, then describe what you want.<br>
            Supports <strong>.docx</strong>, <strong>.pptx</strong>, and <strong>.pdf</strong> sources.<br>
            Or just ask anything — research, writing, code.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Quick-action chips — one click sends a ready-made request
    cols = st.columns(len(SUGGESTIONS) + 1)
    for col, (icon, label, prompt) in zip(cols, SUGGESTIONS):
        with col:
            if st.button(f"{icon} {label}", key=f"sugg_{label}",
                         use_container_width=True):
                handle_input(prompt)
    with cols[-1]:
        if st.button("🎓 Study this doc", key="sugg_learn",
                     use_container_width=True):
            st.session_state.app_mode = "learn"
            rerun()
