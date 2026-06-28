# ======================================================================================
# auth.py
# ======================================================================================
# Module: Authentication, Session Validation & Identity Control Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module manages user authentication and session lifecycle control for
# EduDraft Studio via Supabase Auth.
#
# It provides helper functions for sign-up, login, logout, session validation,
# token rotation handling, and authenticated identity retrieval, while ensuring
# that Row Level Security (RLS) can operate correctly against Supabase-backed data.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. User Sign-Up
#    - Registers new users via Supabase Auth
#    - Handles email-confirmation-required flows
#    - Returns user-friendly messages for onboarding states
#
# 2. User Login
#    - Authenticates users with email and password
#    - Extracts and returns usable session tokens
#    - Normalises login output into app-friendly session dictionaries
#
# 3. User Logout
#    - Signs out the active Supabase session
#    - Clears application-level session state safely
#
# 4. Session Token Handling
#    - Supports both:
#        • dictionary-based session state
#        • Supabase Session object style
#    - Extracts access and refresh tokens consistently
#
# 5. Session Validation & Refresh
#    - Verifies that a session is present and valid
#    - Calls supabase.auth.set_session(...) to rehydrate/refresh auth state
#    - Handles refresh-token rotation safely by storing updated tokens back into session state
#
# 6. Identity Lookup
#    - Retrieves the currently authenticated user
#    - Provides "who am I" diagnostics for app-level visibility
#
# 7. RLS Authentication Support
#    - Sets PostgREST auth headers after session validation
#    - Ensures authenticated database queries work correctly under Supabase RLS policies
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - Supabase client from config.py
# - datetime / timezone utilities
# - typing helpers for structured return contracts
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Session safety first
# - Defensive handling across Supabase SDK response variations
# - Clear app-friendly return values
# - Token rotation awareness to prevent stale-session failures
# - Authentication state should be explicit, validated, and reusable
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module is the identity gatekeeper of the platform.
#
# In practice:
#   - app.py uses it for sign-up, login, and logout flows
#   - rate_limit.py and credits.py rely on validated user identity
#   - any Supabase RLS-protected operation depends on _require_session(...)
#
# It therefore sits upstream of:
#   - wallet access
#   - usage tracking
#   - draft/template persistence
#   - user-specific feature access
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - Token rotation handling is critical:
#     refresh tokens may change after set_session(...)
#     and must be written back into session state
# - Failure to persist rotated tokens can lead to:
#     "Invalid Refresh Token: Already Used"
# - This module should remain stable and conservative, as changes here affect
#   all authenticated workflows in the system
#
# ======================================================================================


from datetime import datetime, timezone
from typing import Any, Dict, Optional, Tuple

from config import supabase

def _session_tokens(session_state):
    """
    Supports BOTH:
      - dict session_state: {"access_token": "...", "refresh_token": "..."}
      - supabase Session object: session_state.access_token / session_state.refresh_token
    Returns: (access_token, refresh_token)
    """
    if not session_state:
        return None, None

    # dict style
    if isinstance(session_state, dict):
        return session_state.get("access_token"), session_state.get("refresh_token")

    # object style
    access = getattr(session_state, "access_token", None)
    refresh = getattr(session_state, "refresh_token", None)
    return access, refresh

def auth_signup(email: str, password: str):
    email = (email or "").strip().lower()
    password = (password or "").strip()
    if not email or not password:
        return "Enter email + password.", None

    try:
        res = supabase.auth.sign_up({"email": email, "password": password})

        # When "Confirm email" is ON, Supabase often creates the user but returns NO session yet.
        # That is normal: user must click the confirmation email first.
        user = getattr(res, "user", None)
        sess = getattr(res, "session", None)

        if user and not sess:
            return (
                "✅ Account created.\n"
                "📩 Please confirm your email (check your inbox), then come back and log in.",
                None
            )

        if user and sess:
            return f"✅ Signed up: {user.email}", sess

        return "Signup created; email confirmation may be required. Check inbox.", None

    except Exception as e:
        return f"Signup failed: {type(e).__name__}: {e}", None


def auth_login(email: str, password: str):
    email = (email or "").strip().lower()
    password = (password or "").strip()
    if not email or not password:
        return "Enter email + password.", None

    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        if res.user is None or res.session is None:
            return "Login failed. Check credentials.", None

        return f"✅ Logged in: {res.user.email}", {
            "access_token": res.session.access_token,
            "refresh_token": res.session.refresh_token,
            "user_id": res.user.id,
            "email": res.user.email,
        }
    except Exception as e:
        return f"Login failed: {type(e).__name__}: {e}", None


def auth_logout(session_state):
    try:
        try:
            supabase.auth.sign_out()
        except Exception:
            pass
        return "✅ Logged out.", None
    except Exception as e:
        return f"Logout failed: {type(e).__name__}: {e}", None


def auth_whoami(session_state):
    if not session_state:
        return "Not logged in."
    try:
        access_token, refresh_token = _session_tokens(session_state)
        if not access_token or not refresh_token:
            return "Session missing tokens. Please log in again."

        supabase.auth.set_session(access_token, refresh_token)
        user = supabase.auth.get_user().user
        if not user:
            return "Session invalid/expired."
        return f"Logged in as:\n- Email: {user.email}\n- UID: {user.id}"
    except Exception as e:
        return f"WhoAmI failed: {type(e).__name__}: {e}"


def _require_session(sess):
    """
    Returns: (access_token, refresh_token, user_id, err)
    Also sets PostgREST auth so RLS works.
    IMPORTANT: If Supabase rotates refresh tokens, we must store the NEW token back into sess,
    otherwise we will get: "Invalid Refresh Token: Already Used"
    """
    if not sess or not isinstance(sess, dict):
        return None, None, None, "❌ Not logged in."

    access_token, refresh_token = _session_tokens(sess)

    if not access_token or not refresh_token:
        return None, None, None, "❌ Session missing tokens. Please log in again."

    try:
        # set_session may rotate tokens; capture and store the new ones
        res = supabase.auth.set_session(access_token, refresh_token)

        # supabase-py responses vary slightly by version; be defensive
        new_sess = None
        if hasattr(res, "session"):
            new_sess = res.session
        elif isinstance(res, dict):
            new_sess = res.get("session")

        if new_sess:
            new_access = getattr(new_sess, "access_token", None) or (new_sess.get("access_token") if isinstance(new_sess, dict) else None)
            new_refresh = getattr(new_sess, "refresh_token", None) or (new_sess.get("refresh_token") if isinstance(new_sess, dict) else None)

            if new_access:
                sess["access_token"] = new_access
                access_token = new_access
            if new_refresh:
                sess["refresh_token"] = new_refresh
                refresh_token = new_refresh

        user = supabase.auth.get_user().user
        if not user:
            return None, None, None, "❌ Session invalid/expired. Please log in again."

        supabase.postgrest.auth(access_token)
        return access_token, refresh_token, user.id, None

    except Exception as e:
        return None, None, None, f"❌ Session check failed: {type(e).__name__}: {e}"




