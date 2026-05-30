"""
compat.py — Streamlit version compatibility helpers.

st.rerun() was added in Streamlit 1.27.0.
st.experimental_rerun() is the old name (still works but deprecated in newer versions).
This shim makes the code work on both old and new Streamlit installs.
"""

import streamlit as st


def rerun() -> None:
    """
    Trigger a Streamlit script rerun.
    Works on Streamlit < 1.27 (experimental_rerun) and >= 1.27 (rerun).
    """
    if hasattr(st, "rerun"):
        st.rerun()
    else:
        st.experimental_rerun()  # type: ignore[attr-defined]
