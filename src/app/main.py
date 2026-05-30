"""
main.py — Streamlit entry point.
Routes to login if not authenticated, else to chat.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from dotenv import load_dotenv
load_dotenv()

import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="DocuForge",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inject CSS
css_path = Path(__file__).parent / "styles" / "main.css"
if css_path.exists():
    with open(css_path) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

from src.auth.session import init_session, is_authenticated

init_session()

if not is_authenticated():
    from src.app.pages.login import render_login
    render_login()
else:
    from src.app.pages.chat import render_chat, render_input, handle_input
    render_chat()
    user_input = render_input()
    if user_input:
        handle_input(user_input)
