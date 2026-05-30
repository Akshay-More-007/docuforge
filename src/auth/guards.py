"""
guards.py — Page-level auth guards.

st.switch_page() only works with multi-page apps using Streamlit's pages/ dir.
For our single-file architecture (main.py handles routing), we just stop execution
here — main.py already shows the login form when not authenticated.
"""

import streamlit as st
from src.auth.session import is_authenticated


def require_auth():
    """
    Stop rendering the chat page if the user is not authenticated.
    main.py already handles showing the login form, so no redirect needed.
    """
    if not is_authenticated():
        st.warning("Please sign in to continue.")
        st.stop()
