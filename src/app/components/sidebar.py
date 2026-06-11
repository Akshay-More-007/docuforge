"""
sidebar.py — Navigation sidebar: brand, new chat, history, upload, logout.
"""

import os
import tempfile
import uuid
import streamlit as st
from src.auth.session import clear_session
from src.auth.supabase_auth import logout
from src.app.compat import rerun


def render_sidebar():
    with st.sidebar:
        # Brand
        st.markdown("""
        <div style="padding: 8px 0 24px 0;">
            <div class="brand">⚡ DocuForge</div>
            <div class="brand-sub">Document Intelligence</div>
        </div>
        """, unsafe_allow_html=True)

        # Mode navigation: Chat ↔ Learn
        mode = st.session_state.get("app_mode", "chat")
        nav1, nav2 = st.columns(2)
        with nav1:
            if st.button("💬 Chat", use_container_width=True,
                         type="primary" if mode == "chat" else "secondary"):
                st.session_state.app_mode = "chat"
                rerun()
        with nav2:
            if st.button("🎓 Learn", use_container_width=True,
                         type="primary" if mode == "learn" else "secondary"):
                st.session_state.app_mode = "learn"
                rerun()

        # New Chat
        if st.button("＋  New Chat", use_container_width=True, type="primary"):
            st.session_state.current_session_id = str(uuid.uuid4())
            st.session_state.messages = []
            st.session_state.source_doc_path = None
            st.session_state.template_doc_path = None
            st.session_state.requirements = []
            st.session_state.last_output_path = None   # ← clear old download
            rerun()

        st.markdown("---")

        # Document upload section
        render_file_uploader()

        st.markdown("---")

        # User info + logout
        email = st.session_state.get("user_email", "")
        st.markdown(
            f"<p style='font-size:11px; color:#666; margin-bottom:8px;'>Signed in as<br>"
            f"<span style='color:#f5a623;'>{email}</span></p>",
            unsafe_allow_html=True,
        )

        if st.button("Sign Out", use_container_width=True):
            logout()
            clear_session()
            rerun()


def render_file_uploader():
    """Sidebar document upload section — supports .docx, .pptx, .pdf."""
    st.markdown(
        "<p style='font-size:11px; color:#888; letter-spacing:1px; text-transform:uppercase; margin-bottom:8px;'>Documents</p>",
        unsafe_allow_html=True,
    )

    source = st.file_uploader(
        "Source Document",
        type=["docx", "pptx", "pdf"],      # ← PDF now supported
        key="source_upload",
        help="Upload the document to reformat (.docx, .pptx, or .pdf)",
    )

    template = st.file_uploader(
        "Template (optional)",
        type=["docx", "pptx"],             # PDF can't be used as a style template
        key="template_upload",
        help="Upload a template to apply styles from (.docx or .pptx)",
    )

    if source:
        path = _save_upload(source, "source")
        if path != st.session_state.get("source_doc_path"):
            st.session_state.source_doc_path = path
            st.session_state.source_doc_name = source.name
        _doc_card("📄", source.name, len(source.getbuffer()))

    if template:
        path = _save_upload(template, "template")
        if path != st.session_state.get("template_doc_path"):
            st.session_state.template_doc_path = path
            st.session_state.template_doc_name = template.name
        _doc_card("📋", template.name, len(template.getbuffer()))

    # Show persistent cards (with the real filename) when no new file is
    # selected this render — the doc stays attached for the whole session.
    if st.session_state.get("source_doc_path") and not source:
        _doc_card("📄", st.session_state.get("source_doc_name", "Source document"))
    if st.session_state.get("template_doc_path") and not template:
        _doc_card("📋", st.session_state.get("template_doc_name", "Template"))

    if st.session_state.get("source_doc_path") or st.session_state.get("template_doc_path"):
        if st.button("✕  Detach documents", use_container_width=True):
            st.session_state.source_doc_path = None
            st.session_state.source_doc_name = None
            st.session_state.template_doc_path = None
            st.session_state.template_doc_name = None
            rerun()


def _doc_card(icon: str, name: str, size_bytes: int | None = None):
    size = ""
    if size_bytes:
        kb = size_bytes / 1024
        size = f"{kb / 1024:.1f} MB" if kb > 1024 else f"{kb:.0f} KB"
    meta = f"<div style='font-size:11px;color:#666;'>{size or 'Attached'}</div>"
    st.markdown(
        f"<div class='doc-card'><div class='doc-name'>{icon} {name}</div>{meta}</div>",
        unsafe_allow_html=True,
    )


def _save_upload(file, prefix: str) -> str:
    """Save uploaded file to a cross-platform temp directory and return path."""
    session_id = st.session_state.get("current_session_id", "default")
    out_dir = os.path.join(tempfile.gettempdir(), "docuforge", session_id)
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"{prefix}_{file.name}")
    with open(path, "wb") as f:
        f.write(file.getbuffer())
    return path
