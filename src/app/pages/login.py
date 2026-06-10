"""
login.py — Login + Signup page.
"""

import streamlit as st
from src.auth.supabase_auth import login, signup
from src.auth.session import set_session
from src.app.compat import rerun


def render_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])

    with col2:
        st.markdown("""
        <div style="text-align:center; padding: 48px 0 32px 0;">
            <div class="brand">DocuForge</div>
            <div class="brand-sub">Document Intelligence</div>
        </div>
        """, unsafe_allow_html=True)

        tab_login, tab_signup = st.tabs(["Sign In", "Sign Up"])

        with tab_login:
            _render_login_form()

        with tab_signup:
            _render_signup_form()


def _render_login_form():
    with st.form("login_form"):
        email = st.text_input("Email", placeholder="you@example.com")
        password = st.text_input("Password", type="password", placeholder="••••••••")
        submitted = st.form_submit_button("Sign In", type="primary", use_container_width=True)

    if submitted:
        if not email or not password:
            st.error("Email and password required.")
            return

        with st.spinner("Signing in..."):
            result = login(email, password)

        if result["success"]:
            set_session(result["user"], result["session"])
            st.success("Signed in!")
            rerun()
        else:
            error = result.get("error", "Login failed.")
            # Clean up Supabase error messages
            if "Invalid login" in error or "invalid_grant" in error:
                st.error("Incorrect email or password.")
            elif "Email not confirmed" in error:
                st.error("Please confirm your email before signing in.")
            else:
                st.error(f"Login failed: {error}")


def _render_signup_form():
    st.markdown(
        "<p style='font-size:12px; color:#888; margin-bottom:12px;'>"
        "Access is invite-only. Only pre-approved emails can create accounts.</p>",
        unsafe_allow_html=True,
    )

    with st.form("signup_form"):
        email = st.text_input("Email", placeholder="you@example.com", key="su_email")
        password = st.text_input("Password", type="password", placeholder="Min 8 characters", key="su_pass")
        confirm = st.text_input("Confirm Password", type="password", placeholder="Repeat password", key="su_confirm")
        submitted = st.form_submit_button("Create Account", use_container_width=True)

    if submitted:
        if not email or not password:
            st.error("All fields required.")
            return
        if password != confirm:
            st.error("Passwords don't match.")
            return
        if len(password) < 8:
            st.error("Password must be at least 8 characters.")
            return

        # Check allowed emails — fail closed: with no whitelist configured,
        # signup is disabled entirely (this app is deployed on the public internet).
        import os
        allowed = os.getenv("ALLOWED_EMAILS", "")
        allowed_list = [e.strip().lower() for e in allowed.split(",") if e.strip()]
        if not allowed_list:
            st.error("Signup is disabled: no approved emails are configured.")
            return
        if email.lower() not in allowed_list:
            st.error("This email is not approved for access.")
            return

        with st.spinner("Creating account..."):
            result = signup(email, password)

        if result["success"]:
            st.success("Account created! Check your email to confirm, then sign in.")
        else:
            error = result.get("error", "Signup failed.")
            st.error(f"Signup failed: {error}")
