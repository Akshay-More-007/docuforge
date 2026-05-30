"""
supabase_auth.py — Signup, login, logout via Supabase Auth.
"""

import os
import logging
from supabase import create_client, Client

logger = logging.getLogger(__name__)

_client: Client | None = None


def get_supabase() -> Client:
    global _client
    if _client is None:
        url = os.getenv("SUPABASE_URL")
        key = os.getenv("SUPABASE_ANON_KEY")
        if not url or not key:
            raise EnvironmentError("SUPABASE_URL and SUPABASE_ANON_KEY must be set in .env")
        _client = create_client(url, key)
    return _client


def login(email: str, password: str) -> dict:
    """
    Returns: {"success": True, "user": ..., "session": ...}
          or {"success": False, "error": "..."}
    """
    try:
        sb = get_supabase()
        res = sb.auth.sign_in_with_password({"email": email, "password": password})
        return {"success": True, "user": res.user, "session": res.session}
    except Exception as e:
        logger.warning(f"[Auth] Login failed for {email}: {e}")
        return {"success": False, "error": str(e)}


def signup(email: str, password: str) -> dict:
    try:
        sb = get_supabase()
        res = sb.auth.sign_up({"email": email, "password": password})
        return {"success": True, "user": res.user}
    except Exception as e:
        logger.warning(f"[Auth] Signup failed for {email}: {e}")
        return {"success": False, "error": str(e)}


def logout():
    try:
        get_supabase().auth.sign_out()
    except Exception as e:
        logger.warning(f"[Auth] Logout error: {e}")


def refresh_session(refresh_token: str) -> dict:
    try:
        sb = get_supabase()
        res = sb.auth.refresh_session(refresh_token)
        return {"success": True, "session": res.session}
    except Exception as e:
        return {"success": False, "error": str(e)}
