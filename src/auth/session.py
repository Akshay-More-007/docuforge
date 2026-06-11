"""
session.py — JWT validation + Streamlit session_state management.
"""

import os
import time
import logging
import streamlit as st
from src.auth.supabase_auth import refresh_session

logger = logging.getLogger(__name__)


def init_session():
    """Initialize all session state keys on first load."""
    defaults = {
        "authenticated": False,
        "user_id": None,
        "user_email": None,
        "access_token": None,
        "refresh_token": None,
        "token_expires_at": None,
        "current_session_id": None,
        "messages": [],
        "source_doc_path": None,
        "template_doc_path": None,
        "requirements": [],
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def set_session(user, session):
    """Store auth data in session_state after successful login."""
    st.session_state.authenticated = True
    st.session_state.user_id = str(user.id)
    st.session_state.user_email = user.email
    st.session_state.access_token = session.access_token
    st.session_state.refresh_token = session.refresh_token
    st.session_state.token_expires_at = session.expires_at


def clear_session():
    """Wipe auth state on logout."""
    keys = [
        "authenticated", "user_id", "user_email",
        "access_token", "refresh_token", "token_expires_at",
        "current_session_id", "messages",
        "source_doc_path", "template_doc_path", "requirements",
    ]
    for key in keys:
        st.session_state[key] = None
    st.session_state.authenticated = False
    st.session_state.messages = []


def is_authenticated() -> bool:
    """Check if user is logged in with a valid (or refreshable) token."""
    # Local development bypass — lets the app run without a live Supabase
    # project. NEVER set DOCUFORGE_DEV_NO_AUTH in a deployed environment.
    if os.getenv("DOCUFORGE_DEV_NO_AUTH") == "1":
        if not st.session_state.get("authenticated"):
            st.session_state.authenticated = True
            st.session_state.user_id = "dev-local"
            st.session_state.user_email = "dev@localhost"
        return True

    if not st.session_state.get("authenticated"):
        return False

    expires_at = st.session_state.get("token_expires_at")
    if expires_at and time.time() > expires_at - 60:
        # Token expiring soon — attempt silent refresh
        refresh_token = st.session_state.get("refresh_token")
        if refresh_token:
            result = refresh_session(refresh_token)
            if result["success"]:
                session = result["session"]
                st.session_state.access_token = session.access_token
                st.session_state.refresh_token = session.refresh_token
                st.session_state.token_expires_at = session.expires_at
                logger.info("[Session] Token silently refreshed")
            else:
                clear_session()
                return False
        else:
            clear_session()
            return False

    return True
