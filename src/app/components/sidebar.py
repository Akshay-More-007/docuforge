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
        st.markdown(
            f"<div class='doc-card'><div class='doc-name'>📄 {source.name}</div>"
            f"<div style='font-size:11px;color:#666;'>Source loaded</div></div>",
            unsafe_allow_html=True,
        )

    if template:
        path = _save_upload(template, "template")
        if path != st.session_state.get("template_doc_path"):
            st.session_state.template_doc_path = path
        st.markdown(
            f"<div class='doc-card'><div class='doc-name'>📋 {template.name}</div>"
            f"<div style='font-size:11px;color:#666;'>Template loaded</div></div>",
            unsafe_allow_html=True,
        )

    # Show persistent labels when no new file is selected this render
    if st.session_state.get("source_doc_path") and not source:
        st.markdown(
            "<div class='doc-card'><div class='doc-name'>📄 Source loaded</div></div>",
            unsafe_allow_html=True,
        )
    if st.session_state.get("template_doc_path") and not template:
        st.markdown(
            "<div class='doc-card'><div class='doc-name'>📋 Template loaded</div></div>",
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
