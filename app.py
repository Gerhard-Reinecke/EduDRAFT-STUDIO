# ======================================================================================
# app.py
# ======================================================================================
# Module: Main Application Orchestrator, UI Composition & Workflow Control Engine
#
# System: EduDRAFT STUDIO (Marike App)
# Version: 4.9
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module is the primary application entry point and orchestration layer of
# EduDRAFT STUDIO.
#
# It composes the Gradio user interface, coordinates authentication and legal-gating
# flows, manages profile and account state, connects generation and export pipelines,
# controls plan-gated feature access, and binds the system's modules into a single
# working platform.
#
# It is the central runtime layer where user actions become structured application
# behaviour.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. UI Composition & Navigation
#    - Builds and manages the Gradio interface
#    - Controls tab visibility, routing, banner state, and interactive UI behaviour
#    - Applies custom CSS and presentation rules for a polished user experience
#
# 2. Authentication & Legal Gatekeeping
#    - Handles sign-up, login, logout, and post-auth routing
#    - Enforces legal acknowledgement requirements before granting app access
#    - Integrates account state checks such as paused/deletion-requested lock conditions
#
# 3. Profile, Account & Avatar Management
#    - Loads and saves user profile data
#    - Manages avatar upload, cropping, default-avatar fallback, and profile banners
#    - Supports admin-only default avatar management and legal-config administration
#
# 4. Draft, Template & Workspace Orchestration
#    - Connects workspace actions to draft generation and editing flows
#    - Coordinates draft loading, renaming, version handling, and template application
#    - Binds template-engine behaviour into the main application workflow
#
# 5. Generation Pipeline Integration
#    - Connects the UI to:
#        • llm.py
#        • ingest_pdf.py
#        • ingest_docx.py
#        • ingest_pptx.py
#        • template_engine.py
#        • diagram_library.py
#        • exports.py
#    - Ensures generation, transformation, and export actions are accessible through UI events
#
# 6. Monetisation, Gating & Usage Controls
#    - Applies Free vs Pro feature restrictions
#    - Surfaces credit status and image-generation cost warnings
#    - Integrates rate limiting, wallet balance checks, and spend enforcement
#
# 7. Export & Preview Coordination
#    - Routes content into DOCX, PPTX, preview, and visual-rendering workflows
#    - Coordinates export confirmation logic where credits or images are involved
#    - Ensures the output experience remains aligned with user permissions and balances
#
# 8. Administrative & Diagnostic Hooks
#    - Supports debug modes such as RUN_DIAGRAM_DEBUG
#    - Exposes admin-facing utilities for legal config and system visibility
#    - Acts as the runtime shell in which diagnostics can be surfaced or tested
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - Gradio → UI framework
# - config.py → system clients, feature flags, limits, categories
# - auth.py → authentication and session management
# - llm.py → prompt and generation pipeline
# - exports.py → export packaging and preview rendering
# - ingest_pdf.py / ingest_docx.py / ingest_pptx.py → file ingestion workflows
# - template_engine.py → donor/template intelligence
# - rate_limit.py → usage controls
# - credits.py → wallet and spend control
# - diagram_library.py → diagram generation and fallback visuals
# - OpenAI / Supabase / pycountry and related runtime libraries
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - One central runtime orchestrator
# - Clear separation between UI logic and service modules
# - Defensive handling of auth, session, and profile state
# - Teacher-first interaction flow
# - Commercial controls integrated without breaking usability
# - Stable tab/state behaviour across login, profile, and plan transitions
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module is the live application shell of EduDRAFT STUDIO.
#
# In practical terms:
#   - config.py prepares the system
#   - service modules provide capabilities
#   - app.py binds everything together into one usable product
#
# It therefore sits above nearly every runtime module and directly governs:
#   - user experience
#   - workflow sequencing
#   - state transitions
#   - gated access
#   - end-to-end feature execution
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This file is operationally critical and highly interconnected
# - Changes here can affect authentication, navigation, generation, exports,
#   monetisation, and user account behaviour simultaneously
# - Debug hooks should remain controlled and explicit
# - This module should be changed carefully, with regression checks across all major tabs
#   and user states
#
# ======================================================================================

import os

# TEMP DEBUG: run diagram debug runner instead of starting Gradio
if os.environ.get("RUN_DIAGRAM_DEBUG", "0") == "1":
    import debug_diagram_library
    debug_diagram_library.run_debug()
    raise SystemExit(0)

import re
import io
import uuid
import json
import tempfile
import subprocess
from datetime import datetime, timezone
import gradio as gr
import pycountry
from typing import Optional
from config import (
    client,
    supabase,
    SUPABASE_DEBUG,
    DAILY_LIMIT_GENERATE,
    DAILY_LIMIT_TRANSCRIBE,
    TEMPLATE_CATEGORIES,
    supabase_secrets_masked,
    supabase_healthcheck)
from auth import (
    auth_signup,
    auth_login,
    auth_logout,
    auth_whoami,
    _require_session)
from llm import (
    safe_err,
    transcribe_audio,
    build_user_request,
    call_llm,
    split_sections,
    combine_doc_and_memo)
from exports import (
    md_to_docx_with_editable_equations,
    outline_to_pptx_with_math,
    build_mathjax_html,
    render_visuals_for_export,
    normalize_math_delimiters_for_pandoc,
    _strip_bundle_sections_for_docx,
    _strip_outer_md_fence)
from ingest_pdf import action_generate_from_pdf, action_analyze_pdf
from ingest_docx import action_generate_from_docx, action_analyze_docx
from ingest_pptx import action_generate_from_pptx, action_analyze_pptx
from template_engine import (
    action_analyze_template_upload,
    save_template_record,
    load_template_bundle_from_db,
    apply_template_to_draft,
    pack_template_bundle,
    handle_teacher_confirmation,
    Decision,
)
from rate_limit import check_rate_limit, get_rate_limit_display
from credits import (
    credits_status_text,
    credits_needed_for_markdown,
    get_balance,
    spend_credits)
from config import ENABLE_IMAGE_GEN
import base64
from openai import OpenAI
import diagram_library


CUSTOM_CSS = """
/* -------------------------------------------------------
   Export area styling (V3.1 FIXED)
   ------------------------------------------------------- */
#action_buttons_row {
  gap: 14px !important;
  background: #000000 !important;
  padding: 2px !important;
  border-radius: 14px !important;
}

#action_buttons_row button,
#action_buttons_row [role="button"] {
  border-radius: 12px !important;
}

/* Export warning */
#export_credit_note {
  background: #C26028 !important;
  border: 1px solid #C26028 !important;
  padding: 12px 14px !important;
  border-radius: 12px !important;
  margin-top: 10px !important;
  box-shadow: 0 6px 18px rgba(0,0,0,0.35) !important;
}

#export_credit_note * {
  color: #FFFFFF !important;
}

/* Checkbox styling */
#export_credit_confirm {
  margin-top: 8px !important;
  padding: 8px 12px !important;
  border-radius: 8px !important;
  background: var(--neutral-700) !important;
}

#export_credit_confirm input[type="checkbox"] {
  accent-color: #C26028 !important;
  transform: scale(1.15);
  margin-right: 8px !important;
}

/* Disabled button */
#export_btn button:disabled,
#export_btn [aria-disabled="true"] {
  opacity: 0.55 !important;
  cursor: not-allowed !important;
}

/* -------------------------------------------------------
   Header / Banner styling (clean + valid)
   ------------------------------------------------------- */

/* Make the GLOBAL banner area roomier */
#global_banner {
  padding: 10px 12px !important;
  border-radius: 14px !important;
}

/* Make the PROFILE banner row taller / roomier */
#profile_banner_row {
  align-items: center !important;
  padding: 10px 12px !important;
  min-height: 150px !important;
  border-radius: 14px !important;
}

/* === AVATAR SIZE CONTROLS ===
   Change AVATAR_PX only (MUST stay square for a circle).
*/

/* Profile avatar circle (outer mask) */
#profile_banner_avatar {
  width: 180px !important;       /* <-- tweak size */
  height: 180px !important;      /* <-- MUST match width for a true circle */
  max-width: 180px !important;
  max-height: 180px !important;
  border-radius: 9999px !important;
  overflow: hidden !important;
}

/* Inner image: full-bleed cover (no rounding here) */
#profile_banner_avatar img {
  width: 100% !important;                 /* <-- IMPORTANT: back to 100% (server crop does zoom) */
  height: 100% !important;
  object-fit: cover !important;
  object-position: center 20% !important; /* <-- tweak up/down */
  border-radius: 0 !important;            /* <-- IMPORTANT: prevent double-rounding */
  display: block !important;
}

/* Global header avatar circle (outer mask) */
#global_avatar_img {
  width: 64px !important;
  height: 64px !important;
  max-width: 64px !important;
  max-height: 64px !important;
  border-radius: 9999px !important;
  overflow: hidden !important;
}

/* Global inner image: ALSO no rounding here */
#global_avatar_img img {
  width: 100% !important;
  height: 100% !important;
  object-fit: cover !important;
  object-position: center 35% !important;
  border-radius: 0 !important;            /* <-- IMPORTANT: prevent double-rounding */
  display: block !important;
}

/* Make the name bigger in the Profile banner */
#profile_banner_row .prose strong,
#profile_banner_row strong {
  font-size: 28px !important;
  line-height: 1.15 !important;
}

/* Hide the "Image" label bar ONLY inside the two header avatars */
#global_avatar_img .label,
#profile_banner_avatar .label,
#global_avatar_img label,
#profile_banner_avatar label {
  display: none !important;
}

/* Hide toolbar/overlay icons ONLY (do NOT nuke all buttons) */
#global_avatar_img .icon-buttons,
#profile_banner_avatar .icon-buttons,
#global_avatar_img .toolbar,
#profile_banner_avatar .toolbar,
#global_avatar_img .overlay,
#profile_banner_avatar .overlay {
  display: none !important;
}

/* =========================
   Legal doc scroll box (B2)
   ========================= */
.legal_doc_box {
  max-height: 300px;
  overflow-y: auto;
  padding: 12px 14px;
  border: 1px solid rgba(255,255,255,0.10);
  border-radius: 12px;
  background: rgba(0,0,0,0.18);
}

/* Optional: make headings inside the legal doc box look tidy */
.legal_doc_box h1, .legal_doc_box h2, .legal_doc_box h3 {
  margin-top: 10px;
}
.legal_doc_box p {
  margin: 8px 0;
}
"""


# ============================
# LEGAL DOC VERSIONS (single source of truth)
# ============================
LEGAL_VERSIONS = {
    "privacy": "PP-2026-02-18",
    "terms":   "TOU-2026-02-18",
    "tcs":     "TCS-2026-02-18",
}

# ============================
# LEGAL GATE (Login/Signup)
# ============================

LEGAL_DOCS_GATE = {
    # doc_key: (profiles_ack_at_col, profiles_version_col, fallback_version, legal_ack_events.doc_type)
    "privacy": ("privacy_policy_ack_at", "privacy_policy_version", LEGAL_VERSIONS["privacy"], "privacy_policy"),
    "terms":   ("terms_of_use_ack_at", "terms_of_use_version", LEGAL_VERSIONS["terms"], "terms_of_use"),
    "tcs":     ("terms_and_conditions_ack_at", "terms_and_conditions_version", LEGAL_VERSIONS["tcs"], "terms_and_conditions"),
}


def _get_required_legal_versions(sb):
    """
    Returns dict: {doc_key: required_version}
    Source of truth: legal_config table (doc_type + current_version)
    Fallback: LEGAL_VERSIONS
    """
    required = {
        "privacy": LEGAL_VERSIONS["privacy"],
        "terms":   LEGAL_VERSIONS["terms"],
        "tcs":     LEGAL_VERSIONS["tcs"],
    }

    try:
        rows = sb.table("legal_config").select("doc_type,current_version").execute()
        for r in (rows.data or []):
            k = (r.get("doc_type") or "").strip().lower()
            v = (r.get("current_version") or "").strip()
            if k in required and v:
                required[k] = v
    except Exception:
        pass

    return required


def _admin_legal_fetch(sess):
    """
    Admin: show current required versions from legal_config (DB truth).
    """
    if not sess:
        return "❌ Not signed in."

    try:
        sb = _sb_authed_from_session(sess)

        rows = (
            sb.table("legal_config")
            .select("doc_type,current_version,updated_at,updated_by")
            .order("doc_type")
            .execute()
        )
        data = rows.data or []
        if not data:
            return "⚠️ No rows found in public.legal_config"

        lines = ["### 📜 Legal Config (Required Versions)\n"]
        for r in data:
            dt = (r.get("doc_type") or "").strip()
            ver = (r.get("current_version") or "").strip()
            ua = r.get("updated_at") or ""
            ub = r.get("updated_by") or "—"
            lines.append(f"- **{dt}** → `{ver}`  \n  _updated: {ua} • by: {ub}_")

        return "\n".join(lines)

    except Exception as e:
        return f"❌ Failed to fetch legal_config: {type(e).__name__}: {e}"


def _admin_legal_update(sess, privacy_ver, terms_ver, tcs_ver):
    """
    Admin: update legal_config current_version for privacy/terms/tcs.
    RLS enforces admin-only UPDATE. Trigger stamps updated_at/updated_by.
    """
    if not sess:
        return "❌ Not signed in."

    privacy_ver = (privacy_ver or "").strip()
    terms_ver   = (terms_ver or "").strip()
    tcs_ver     = (tcs_ver or "").strip()

    if not (privacy_ver and terms_ver and tcs_ver):
        return "⚠️ Please fill in all three version fields."

    try:
        sb = _sb_authed_from_session(sess)

        sb.table("legal_config").update({"current_version": privacy_ver}).eq("doc_type", "privacy").execute()
        sb.table("legal_config").update({"current_version": terms_ver}).eq("doc_type", "terms").execute()
        sb.table("legal_config").update({"current_version": tcs_ver}).eq("doc_type", "tcs").execute()

        return "✅ Saved. Version changes will force re-ack for users."

    except Exception as e:
        return f"❌ Save failed: {type(e).__name__}: {e}"


def _legal_ok_for_user(sb, uid: str) -> bool:
    """
    True only if user has acknowledged the CURRENT required versions.
    """
    required = _get_required_legal_versions(sb)

    try:
        prof = (
            sb.table("profiles")
            .select(
                "privacy_policy_ack_at,privacy_policy_version,"
                "terms_of_use_ack_at,terms_of_use_version,"
                "terms_and_conditions_ack_at,terms_and_conditions_version"
            )
            .eq("user_id", uid)
            .limit(1)
            .execute()
        )
        prow = (prof.data[0] if prof.data else {}) or {}
    except Exception:
        prow = {}

    def ok(doc_key: str, ack_col: str, ver_col: str):
        ack_at = prow.get(ack_col)
        ver = (prow.get(ver_col) or "").strip()
        return bool(ack_at) and (ver == (required.get(doc_key) or ""))

    return (
        ok("privacy", "privacy_policy_ack_at", "privacy_policy_version")
        and ok("terms", "terms_of_use_ack_at", "terms_of_use_version")
        and ok("tcs", "terms_and_conditions_ack_at", "terms_and_conditions_version")
    )


def _record_legal_acks_for_user(sess, pp_checked: bool, terms_checked: bool, tcs_checked: bool):
    """
    Writes profile ack + audit rows using current required versions.
    This is called AFTER a successful signup/login (so we already have a session).
    """
    access_token, refresh_token, uid, err = _require_session(sess)
    if err:
        return False, f"❌ {err}"

    sb = _sb_authed_from_session(sess)
    required = _get_required_legal_versions(sb)

    now_iso = _utc_now_iso()

    to_write = []
    if pp_checked:
        to_write.append("privacy")
    if terms_checked:
        to_write.append("terms")
    if tcs_checked:
        to_write.append("tcs")

    # Must have all three checked when we are recording
    if set(to_write) != {"privacy", "terms", "tcs"}:
        return False, "⚠️ Please tick all three legal checkboxes to continue."

    try:
        # Update profiles in one update
        payload = {}
        for k in ("privacy", "terms", "tcs"):
            ack_col, ver_col, fallback_ver, doc_type = LEGAL_DOCS_GATE[k]
            ver = (required.get(k) or fallback_ver or "").strip()
            payload[ack_col] = now_iso
            payload[ver_col] = ver

        sb.table("profiles").update(payload).eq("user_id", uid).execute()

        # Insert audit rows (best-effort)
        for k in ("privacy", "terms", "tcs"):
            ack_col, ver_col, fallback_ver, doc_type = LEGAL_DOCS_GATE[k]
            ver = (required.get(k) or fallback_ver or "").strip()
            try:
                sb.table("legal_ack_events").insert({
                    "user_id": uid,
                    "doc_type": doc_type,
                    "doc_version": ver,
                    "ack_at": now_iso,
                    "client_meta": {"surface": "login_tab"}
                }).execute()
            except Exception:
                pass

        return True, "✅ Legal acknowledgements recorded."

    except Exception as e:
        return False, f"❌ Failed to record legal acknowledgements: {type(e).__name__}: {e}"


def auth_signup_with_legal(email: str, password: str, pp: bool, terms: bool, tcs: bool):
    """
    1) Require checkboxes
    2) Signup
    3) Record legal ack (profiles + audit)
    Only returns ✅ + session if everything succeeded.
    """
    if not (pp and terms and tcs):
        return "⚠️ Please tick all three legal checkboxes to sign up.", None

    msg, sess = auth_signup(email, password)
    if not sess or "✅" not in (msg or ""):
        return msg, None

    ok, ack_msg = _record_legal_acks_for_user(sess, pp, terms, tcs)
    if not ok:
        return ack_msg, None

    return f"{msg}\n{ack_msg}", sess


def auth_login_with_legal(email: str, password: str, pp: bool, terms: bool, tcs: bool):
    """
    Log in only if:
      1) Email + password are correct
      2) Email is confirmed (user clicked the link)
      3) Legal is current (or we record legal now)
    Returns: (status_message, session_dict_or_none)
    """
    email = (email or "").strip().lower()
    password = (password or "").strip()
    if not email or not password:
        return "Enter email + password.", None

    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})

        user = res.user
        sess = res.session

        # If Supabase didn't give us a session, treat as not logged in
        if not user or not sess:
            return "❌ Login failed. Please check your email + password.", None

        # ✅ HARD RULE: must confirm email first
        email_confirmed = bool(getattr(user, "email_confirmed_at", None) or getattr(user, "confirmed_at", None))
        if not email_confirmed:
            try:
                supabase.auth.sign_out()
            except Exception:
                pass
            return "📩 Please confirm your email first (check your inbox), then come back and log in.", None

        # Convert Supabase Session object to the canonical dict format used everywhere else
        session_dict = {
            "access_token": sess.access_token,
            "refresh_token": sess.refresh_token,
            "user_id": user.id,
            "email": user.email,
        }

        # Ensure every logged-in user has a profile row immediately
        ok_profile, profile_msg = ensure_profile_row(session_dict)
        if not ok_profile:
            return f"❌ Login failed while creating profile row: {profile_msg}", None

        # If legal already ok, allow straight through
        uid = getattr(user, "id", None)
        if uid and _legal_ok_for_user(_sb_authed_from_session(session_dict), uid):
            return f"✅ Logged in: {user.email}\n✅ Legal acknowledgements recorded.", session_dict

        # Legal is missing or outdated.
        # Do NOT auto-record legal on login.
        # Return the valid session and let _tab_after_auth(...) route the user to Profile.
        return (
            "⚠️ Your legal acknowledgements are missing or out of date. "
            "Please review and accept the current Privacy Policy, Terms of Use, and T&Cs in your Profile tab.",
            session_dict
        )

    except Exception as e:
        return f"❌ Login failed: {type(e).__name__}: {e}", None


def debug_legal_state(sess):
    try:
        access_token, _, user_id, err = _require_session(sess)
        if err:
            return f"Session error: {err}"

        sb = _sb_authed_from_session(sess)

        res = sb.table("profiles").select("*").eq("user_id", user_id).limit(1).execute()
        row = (res.data or [{}])[0]

        cfg_rows = sb.table("legal_config").select("doc_type,current_version").execute()
        cfg = {r["doc_type"]: r["current_version"] for r in (cfg_rows.data or [])}

        lines = []
        lines.append("=== STORED VERSIONS ===")
        lines.append(f"privacy: {row.get('privacy_policy_version')}")
        lines.append(f"terms:   {row.get('terms_of_use_version')}")
        lines.append(f"tcs:     {row.get('terms_and_conditions_version')}")

        lines.append("\n=== STORED TIMESTAMPS ===")
        lines.append(f"privacy: {row.get('privacy_policy_ack_at')}")
        lines.append(f"terms:   {row.get('terms_of_use_ack_at')}")
        lines.append(f"tcs:     {row.get('terms_and_conditions_ack_at')}")

        lines.append("\n=== REQUIRED VERSIONS ===")
        lines.append(str(cfg))

        is_ok = _legal_ok_for_user(sb, user_id)
        lines.append(f"\n=== FINAL CHECK ===")
        lines.append(f"_legal_ok_for_user: {is_ok}")

        return "\n".join(lines)

    except Exception as e:
        return f"ERROR: {type(e).__name__}: {e}"


# =============================
# CANONICAL HELPERS (ONE COPY ONLY)
# =============================

def _needs_notifications_gate(session_state) -> bool:
    """
    True if the user is authenticated but still blocked by required legal acknowledgements.
    Fail closed: if we cannot verify, keep the gate active.
    """
    if not session_state:
        return False

    try:
        sb = _sb_authed_from_session(session_state)
        uid = _uid_from_session(session_state)
        if not uid:
            return True
        return not _legal_ok_for_user(sb, uid)
    except Exception:
        return True


def _tab_after_auth(auth_message: str, session_state):
    """
    Post-auth routing:
      - no session -> Login
      - authenticated + legal clear -> Home
      - authenticated + legal blocked -> Notifications
    """
    authed = bool(session_state)

    if not authed:
        return ("Login", False)

    if _needs_notifications_gate(session_state):
        return ("Notifications", False)

    return ("Home", True)


def _tab_after_auth_or_stay(fallback_tab_id: str, auth_message: str, session_state):
    """
    If auth failed (no session), stay on the current tab instead of bouncing to sorter.
    If auth succeeded, use the normal routing logic.
    """
    if not session_state:
        return (fallback_tab_id, False)
    return _tab_after_auth(auth_message, session_state)


# =============================
# BROWSER SESSION RESTORE (Gradio fallback)
# =============================
BROWSER_SESSION_STORAGE_KEY = "edudraft_supabase_session_v1"


def _normalize_browser_session_blob(blob):
    """
    Accepts a browser-stored JSON string (or dict) and returns the canonical
    session dict shape used by this app, or None if invalid.
    """
    if not blob:
        return None

    try:
        data = json.loads(blob) if isinstance(blob, str) else blob
    except Exception:
        return None

    if not isinstance(data, dict):
        return None

    access_token = (data.get("access_token") or "").strip()
    refresh_token = (data.get("refresh_token") or "").strip()
    user_id = (data.get("user_id") or "").strip()
    email = (data.get("email") or "").strip()

    if not access_token or not refresh_token or not user_id:
        return None

    return {
        "access_token": access_token,
        "refresh_token": refresh_token,
        "user_id": user_id,
        "email": email,
    }


def _restore_browser_session(browser_blob):
    """
    Restore a previously saved browser session into Gradio state on page load.

    Returns:
      (session_dict_or_none, tab_id, is_logged_in_bool)
    """
    sess = _normalize_browser_session_blob(browser_blob)
    if not sess:
        return None, "Login", False

    try:
        # Validate that the restored session is still usable in the app.
        ok_profile, _profile_msg = ensure_profile_row(sess)
        if not ok_profile:
            return None, "Login", False

        tab_id, logged_in = _tab_after_auth("browser_restore", sess)
        return sess, tab_id, logged_in

    except Exception:
        return None, "Login", False


def _restore_browser_session_with_debug(browser_blob):
    """
    Combined page-load restore:
    - restores session into Gradio state
    - also returns a human-readable debug string
    """
    debug_text = _describe_browser_session_blob(browser_blob)
    sess, tab_id, logged_in = _restore_browser_session(browser_blob)
    return sess, tab_id, logged_in, debug_text


def _describe_browser_session_blob(browser_blob):
    """
    Debug helper: describes what the browser handed back on page load
    or via the manual debug button.
    """
    if not browser_blob:
        return "Browser blob: EMPTY"

    try:
        data = json.loads(browser_blob) if isinstance(browser_blob, str) else browser_blob
    except Exception as e:
        return f"Browser blob: INVALID JSON ({type(e).__name__}: {e})"

    if not isinstance(data, dict):
        return f"Browser blob: NOT A DICT ({type(data).__name__})"

    keys = sorted(list(data.keys()))
    access = bool((data.get("access_token") or "").strip())
    refresh = bool((data.get("refresh_token") or "").strip())
    user_id = bool((data.get("user_id") or "").strip())
    email = (data.get("email") or "").strip()

    return (
        f"Browser blob: PRESENT\n"
        f"keys={keys}\n"
        f"access_token={access}\n"
        f"refresh_token={refresh}\n"
        f"user_id={user_id}\n"
        f"email={email or '(blank)'}"
    )


def _session_to_browser_blob(sess):
    """
    Convert the live session dict into a JSON string that can be safely
    handed to browser-side JS for localStorage persistence.
    """
    if not sess or not isinstance(sess, dict):
        return ""

    access_token = (sess.get("access_token") or "").strip()
    refresh_token = (sess.get("refresh_token") or "").strip()
    user_id = (sess.get("user_id") or "").strip()
    email = (sess.get("email") or "").strip()

    if not access_token or not refresh_token or not user_id:
        return ""

    return json.dumps({
        "access_token": access_token,
        "refresh_token": refresh_token,
        "user_id": user_id,
        "email": email,
    })


def _parse_draft_id(label: str) -> str:
    """
    Label format:
      "<title>  —  <subject>  •  <dt>  | <draft_id>"
    Returns the draft_id.
    """
    if not label:
        return ""
    parts = label.split("|")
    return parts[-1].strip() if parts else ""


def _parse_template_id(choice: str) -> str:
    """
    Dropdown format from list_my_templates:
      "Template Name | full-uuid-here"
    Returns the uuid part.
    """
    if not choice:
        return ""
    if isinstance(choice, list):
        choice = choice[0] if choice else ""
    if "|" in choice:
        return choice.split("|", 1)[1].strip()
    return choice.strip()


def _strip_visual_box_html(md: str) -> str:
    """
    Removes the old V2.4/V2.5 preview HTML boxes like:
      <div style="border:1px dashed ..."> ... </div>
    so they don't pollute saved drafts or loaded versions.
    """
    if not md:
        return md

    # Remove the whole visual placeholder div block
    md = re.sub(
        r'<div\s+style="border:1px dashed[^"]*">.*?</div>\s*',
        '',
        md,
        flags=re.DOTALL | re.IGNORECASE
    )
    return md


def _spend_credits_for_session(sess, amount: float, reason: str, meta: dict | None = None) -> tuple[bool, str, float]:
    _access, _refresh, user_id, err = _require_session(sess)
    if err:
        return False, err, 0.0

    sb = _sb_authed_from_session(sess)
    return spend_credits(user_id, amount, reason, meta, sb=sb)


def _draft_title_from_choice(choice: str) -> str:
    if not choice:
        return ""
    left = choice.split("|", 1)[0].strip()
    title = left.split("—", 1)[0].strip()
    return title


def _banner_md_from_choice(choice: str) -> str:
    t = _draft_title_from_choice(choice)
    if not t:
        return "### 📄 Selected draft: *(none loaded yet)*"
    return f"### 📄 Working on: **{t}**"

def _subject_from_draft_choice(choice: str) -> str:
    """
    Choice format:
      "<title>  —  <subject>  •  <dt>  | <draft_id>"

    Returns a cleaned subject only.
    BS5 RULE:
    - do not let display metadata such as dates leak into subject state
    """
    if not choice:
        return ""

    left = choice.split("|", 1)[0].strip()  # "<title> — <subject> • <dt>"
    if "—" not in left:
        return ""

    subj_part = left.split("—", 1)[1].strip()  # "<subject> • <dt>"
    subj = subj_part.split("•", 1)[0].strip()

    if subj in {"—", ""}:
        return ""

    # Strip trailing date-like suffixes in parentheses, e.g.:
    # "Biology (31 Jan 2026)" -> "Biology"
    subj = re.sub(
        r"\s*\(\s*\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\s*\)\s*$",
        "",
        subj
    ).strip()

    return subj


def _set_subject_state_from_choice(choice: str):
    subj = _subject_from_draft_choice(choice)
    # Update both the hidden box AND the state used on save
    return subj, gr.update(value=subj)


# Gradio-safe UI helpers (no accordion open/close toggles)
def _goto_tab(tab_name: str):
    return gr.update(selected=tab_name)


def _show_hide(show: bool):
    return gr.update(visible=bool(show))


def _sync_password_from_masked(v):
    v = v or ""
    return gr.update(value=v), gr.update(value=v)

def _sync_password_from_plain(v):
    v = v or ""
    return gr.update(value=v), gr.update(value=v)

def _toggle_password_pair(masked_val, plain_val, showing):
    current = plain_val if showing else masked_val
    current = current or ""
    showing = not bool(showing)

    if showing:
        return (
            gr.update(value=current, visible=False),  # masked box
            gr.update(value=current, visible=True),   # plain box
            gr.update(value="🙈"),                    # toggle button
            True                                      # showing state
        )

    return (
        gr.update(value=current, visible=True),       # masked box
        gr.update(value=current, visible=False),      # plain box
        gr.update(value="👁"),                        # toggle button
        False                                         # showing state
    )


def supa_rename_draft(sess, draft_choice, new_title: str):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return err

    draft_id = _parse_draft_id(draft_choice)
    if not draft_id:
        return "No draft selected."

    new_title = (new_title or "").strip()
    if not new_title:
        return "Please enter a new draft name."

    try:
        supabase.postgrest.auth(access_token)
        now = datetime.now(timezone.utc).isoformat()

        (
            supabase.table("drafts")
            .update({"title": new_title, "updated_at": now})
            .eq("id", draft_id)
            .eq("user_id", user_id)
            .execute()
        )

        return f"✅ Renamed draft to: {new_title}"

    except Exception as e:
        return f"❌ Rename failed: {type(e).__name__}: {e}"

def action_image_gen_selfcheck():
    """
    Quick sanity check:
    - prints ENABLE_IMAGE_GEN / model / size
    - makes a real /v1/images call
    - returns (status_text, png_path_or_none)
    """
    enabled = os.environ.get("ENABLE_IMAGE_GEN", "0")
    model = os.environ.get("OPENAI_IMAGE_MODEL", "")
    size = os.environ.get("OPENAI_IMAGE_SIZE", "")
    has_key = bool(os.environ.get("OPENAI_API_KEY", "").strip())

    # Basic config report
    report = [
        f"ENABLE_IMAGE_GEN = {enabled}",
        f"OPENAI_IMAGE_MODEL = {model or '(missing)'}",
        f"OPENAI_IMAGE_SIZE  = {size or '(missing)'}",
        f"OPENAI_API_KEY     = {'✅ set' if has_key else '❌ missing'}",
    ]

    if enabled.strip().lower() not in {"1", "true", "yes", "on"}:
        return "❌ Image-gen is OFF.\n\n" + "\n".join(report), None

    if not has_key:
        return "❌ OPENAI_API_KEY is missing.\n\n" + "\n".join(report), None

    if not model:
        return "❌ OPENAI_IMAGE_MODEL is missing.\n\n" + "\n".join(report), None

    if not size:
        return "❌ OPENAI_IMAGE_SIZE is missing.\n\n" + "\n".join(report), None

    try:
        client = OpenAI()

        prompt = (
            "Black-and-white worksheet line-art diagram of a simple plant cell. "
            "Clean outline, no shading, no color, no labels, white background."
        )

        img = client.images.generate(
            model=model,
            prompt=prompt,
            n=1,
            size=size)

        png_bytes = base64.b64decode(img.data[0].b64_json)

        out = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        with open(out.name, "wb") as f:
            f.write(png_bytes)

        return "✅ Image-gen OK.\n\n" + "\n".join(report), out.name

    except Exception as e:
        return (
            "❌ Image-gen FAILED.\n\n"
            + "\n".join(report)
            + f"\n\nError: {type(e).__name__}: {e}"
        ), None


def supa_get_profile(sess):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return None, f"❌ {err}"

    try:
        supabase.postgrest.auth(access_token)
        res = supabase.table("profiles").select("*").eq("user_id", user_id).limit(1).execute()
        row = res.data[0] if res.data else None
        return row, "✅ Profile loaded."
    except Exception as e:
        return None, f"❌ Failed to load profile: {type(e).__name__}: {e}"


def _account_lock_status(sess):
    """
    Hard lock rules:
      - is_paused == True  -> locked
      - deletion_status in ("requested","processing","deleted") -> locked
    Returns: (locked: bool, reason: str)
    """
    if not sess:
        return False, ""

    try:
        ensure_profile_row(sess)
        prow, _msg = supa_get_profile(sess)
        prow = prow or {}

        if bool(prow.get("is_paused", False)) is True:
            return True, "⏸️ Account Paused"

        st = (prow.get("deletion_status") or "").strip().lower()
        if st in ("requested", "processing", "deleted"):
            return True, "🗑️ Delete Requested"

        return False, ""
    except Exception:
        # Never crash UI because of lock check
        return False, ""

def _is_pro_user(sess) -> bool:
    """
    Returns True if the signed-in user is Pro (paid/premium).
    Uses the same profile fields as supa_plan_badge():
      - profiles.plan OR profiles.subscription_plan
      - profiles.is_pro
    Safe default: False.
    """
    if not sess:
        return False

    try:
        ensure_profile_row(sess)
        prow, _msg = supa_get_profile(sess)
        prow = prow or {}

        plan = (prow.get("plan") or prow.get("subscription_plan") or "").strip().lower()
        is_pro = bool(prow.get("is_pro", False))

        return (plan in ("pro", "paid", "premium")) or is_pro
    except Exception:
        return False

def _pro_button_visibility(sess):
    """
    Profile subscription button visibility:
    - Free user  -> show Activate, hide Downgrade
    - Pro user   -> hide Activate, show Downgrade
    Safe default when not signed in: show Activate, hide Downgrade
    """
    is_pro = _is_pro_user(sess)
    return (
        gr.update(visible=not is_pro),  # activate_pro_btn
        gr.update(visible=is_pro),      # downgrade_pro_btn
    )

def _is_admin_user(sess) -> bool:
    """
    Returns True only if the signed-in user's profile row says is_admin = True.
    Safe default: False.
    """
    if not sess:
        return False

    try:
        ensure_profile_row(sess)
        prow, _msg = supa_get_profile(sess)
        prow = prow or {}
        return bool(prow.get("is_admin", False))
    except Exception:
        return False

def _public_vs_app_tab_visibility(sess):
    """
    Public-only screens when logged out:
      - Login (landing)
      - Sign up
      - Log in

    If logged in and legally blocked:
      - show Notifications only

    If logged in and legally clear:
      - show the normal app tabs
    """
    logged_in = bool(sess and isinstance(sess, dict) and sess.get("access_token"))

    if not logged_in:
        return (
            gr.update(visible=True),   # login_tab
            gr.update(visible=False),  # home_tab
            gr.update(visible=True),   # signup_tab
            gr.update(visible=True),   # login2_tab
            gr.update(visible=False),  # notifications_tab
            gr.update(visible=False),  # workspace_tab
            gr.update(visible=False),  # drafts_tab
            gr.update(visible=False),  # templates_tab
            gr.update(visible=False),  # profile_tab
        )

    gate_active = _needs_notifications_gate(sess)

    if gate_active:
        return (
            gr.update(visible=False),  # login_tab
            gr.update(visible=False),  # home_tab
            gr.update(visible=False),  # signup_tab
            gr.update(visible=False),  # login2_tab
            gr.update(visible=True),   # notifications_tab
            gr.update(visible=False),  # workspace_tab
            gr.update(visible=False),  # drafts_tab
            gr.update(visible=False),  # templates_tab
            gr.update(visible=False),  # profile_tab
        )

    return (
        gr.update(visible=False),  # login_tab
        gr.update(visible=True),   # home_tab
        gr.update(visible=False),  # signup_tab
        gr.update(visible=False),  # login2_tab
        gr.update(visible=False),  # notifications_tab
        gr.update(visible=True),   # workspace_tab
        gr.update(visible=True),   # drafts_tab
        gr.update(visible=True),   # templates_tab
        gr.update(visible=True),   # profile_tab
    )


def _notifications_legal_ui_snapshot(sess):
    """
    Returns:
      privacy_status, terms_status, tcs_status,
      privacy_acc_update, terms_acc_update, tcs_acc_update

    Rule:
      - outstanding documents open automatically
      - current/acknowledged documents stay collapsed
    """
    default_status = "❌ Not acknowledged for current version."
    default_open = gr.update(open=True)

    if not sess:
        return (
            default_status,
            default_status,
            default_status,
            default_open,
            default_open,
            default_open,
        )

    try:
        access_token, refresh_token, user_id, err = _require_session(sess)
        if err:
            return (
                default_status,
                default_status,
                default_status,
                default_open,
                default_open,
                default_open,
            )

        sb = _sb_authed_from_session(sess)
        required = _get_required_legal_versions(sb)

        res = (
            sb.table("profiles")
            .select(
                "privacy_policy_ack_at,privacy_policy_version,"
                "terms_of_use_ack_at,terms_of_use_version,"
                "terms_and_conditions_ack_at,terms_and_conditions_version"
            )
            .eq("user_id", user_id)
            .limit(1)
            .execute()
        )

        row = (res.data[0] if res.data else {}) or {}

        def _one(ack_at, stored_ver, required_ver):
            ack_at = (ack_at or "").strip()
            stored_ver = (stored_ver or "").strip()
            required_ver = (required_ver or "").strip()

            is_current = bool(ack_at) and (stored_ver == required_ver)

            if is_current:
                return (
                    f"✅ Acknowledged on **{ack_at}** ({required_ver})",
                    gr.update(open=False)
                )

            return (
                f"❌ Not acknowledged for current version ({required_ver}).",
                gr.update(open=True)
            )

        privacy_status, privacy_open = _one(
            row.get("privacy_policy_ack_at"),
            row.get("privacy_policy_version"),
            required.get("privacy", "")
        )

        terms_status, terms_open = _one(
            row.get("terms_of_use_ack_at"),
            row.get("terms_of_use_version"),
            required.get("terms", "")
        )

        tcs_status, tcs_open = _one(
            row.get("terms_and_conditions_ack_at"),
            row.get("terms_and_conditions_version"),
            required.get("tcs", "")
        )

        return (
            privacy_status,
            terms_status,
            tcs_status,
            privacy_open,
            terms_open,
            tcs_open,
        )

    except Exception:
        return (
            default_status,
            default_status,
            default_status,
            default_open,
            default_open,
            default_open,
        )


def _notifications_ack_refresh(sess, doc_key: str, checked: bool):
    """
    User-driven legal ack inside Notifications gate.
    Refreshes all three legal status lines and accordion states,
    then decides whether the Notifications gate should remain active
    or route the user to Home.
    """
    _status_text, cb_update = action_ack_legal_doc(sess, doc_key, checked)

    (
        privacy_status,
        terms_status,
        tcs_status,
        privacy_open,
        terms_open,
        tcs_open,
    ) = _notifications_legal_ui_snapshot(sess)

    if _needs_notifications_gate(sess):
        return (
            privacy_status,
            terms_status,
            tcs_status,
            privacy_open,
            terms_open,
            tcs_open,
            cb_update,
            "Notifications",
            False,
        )

    return (
        privacy_status,
        terms_status,
        tcs_status,
        privacy_open,
        terms_open,
        tcs_open,
        cb_update,
        "Home",
        True,
    )


def _profile_legal_status_snapshot(sess):
    """
    Returns the three Profile-tab legal status strings:
      privacy_ack_status, terms_ack_status, tcs_ack_status

    A document counts as current only if:
      - ack timestamp exists
      - stored version matches current required version
    """
    default_status = "❌ Not acknowledged yet."

    if not sess:
        return (default_status, default_status, default_status)

    try:
        access_token, refresh_token, user_id, err = _require_session(sess)
        if err:
            return (default_status, default_status, default_status)

        sb = _sb_authed_from_session(sess)
        required = _get_required_legal_versions(sb)

        res = (
            sb.table("profiles")
            .select(
                "privacy_policy_ack_at,privacy_policy_version,"
                "terms_of_use_ack_at,terms_of_use_version,"
                "terms_and_conditions_ack_at,terms_and_conditions_version"
            )
            .eq("user_id", user_id)
            .limit(1)
            .execute()
        )

        row = (res.data[0] if res.data else {}) or {}

        def _fmt(ack_at, stored_ver, required_ver):
            ack_at = (ack_at or "").strip()
            stored_ver = (stored_ver or "").strip()
            required_ver = (required_ver or "").strip()

            if ack_at and stored_ver == required_ver:
                return f"✅ Acknowledged on **{ack_at}** ({required_ver})"

            return "❌ Not acknowledged yet."

        return (
            _fmt(
                row.get("privacy_policy_ack_at"),
                row.get("privacy_policy_version"),
                required.get("privacy", "")
            ),
            _fmt(
                row.get("terms_of_use_ack_at"),
                row.get("terms_of_use_version"),
                required.get("terms", "")
            ),
            _fmt(
                row.get("terms_and_conditions_ack_at"),
                row.get("terms_and_conditions_version"),
                required.get("tcs", "")
            ),
        )

    except Exception:
        return (default_status, default_status, default_status)


def _admin_tab_visibility(sess):
    return gr.update(visible=_is_admin_user(sess))


def _gate_reason_for_ui(sess) -> str:
    """
    Returns "" if allowed.
    Otherwise returns a human-readable reason string (locked / not logged in / free plan).
    """
    # 1) Hard lock: paused OR delete requested
    locked, reason = _account_lock_status(sess)
    if locked:
        return reason or "Account locked"

    # 2) Not logged in
    if not sess:
        return "🔐 Please log in to access this."

    # 3) Plan gate
    if not _is_pro_user(sess):
        return "🔒 Pro required — this area is locked on the Free plan."

    return ""


def _apply_drafts_plan_lock(sess):
    """
    Outputs MUST match the Drafts tab wiring order:
    (lock_note, draft_search, refresh_btn, drafts_dd, versions_dd, load_btn,
     rename_new_title, rename_btn, delete_version_btn, delete_draft_btn, library_status,
     template_name, template_desc, template_category, share_template,
     save_template_confirm, save_template_status)
    """
    reason = _gate_reason_for_ui(sess)
    locked = bool(reason)

    note_md = (
        f"### {reason}\n\n"
        "Upgrade to Pro to access your Draft Library and reusable content tools."
    ) if locked else ""

    return (
        gr.update(value=note_md, visible=locked),
        gr.update(interactive=not locked),  # draft_search
        gr.update(interactive=not locked),  # refresh_btn
        gr.update(interactive=not locked),  # drafts_dd
        gr.update(interactive=not locked),  # versions_dd
        gr.update(interactive=not locked),  # load_btn
        gr.update(interactive=not locked),  # rename_new_title
        gr.update(interactive=not locked),  # rename_btn
        gr.update(interactive=not locked),  # delete_version_btn
        gr.update(interactive=not locked),  # delete_draft_btn
        (reason if locked else ""),         # library_status
        gr.update(interactive=not locked),  # template_name
        gr.update(interactive=not locked),  # template_desc
        gr.update(interactive=not locked),  # template_category
        gr.update(interactive=not locked),  # share_template
        gr.update(interactive=not locked),  # save_template_confirm
        (reason if locked else "")          # save_template_status
    )


def _apply_templates_plan_lock(sess):
    """
    Outputs MUST match the Templates tab wiring order:
    (lock_note,
     upload_template_confirm, save_analyzed_template_btn, refresh_templates_btn, apply_template_btn,
     delete_template_btn, ref_attach_btn, ref_load_btn,
     templates_dropdown, ref_template_dropdown, up_file,
     load_template_status, generated_template_file)
    """
    reason = _gate_reason_for_ui(sess)
    locked = bool(reason)

    note_md = (
        f"### {reason}\n\n"
        "Upgrade to Pro to use Templates (upload, apply, manage, delete, attach PDFs)."
    ) if locked else ""

    return (
        gr.update(value=note_md, visible=locked),
        gr.update(interactive=not locked),  # upload_template_confirm
        gr.update(interactive=not locked),  # save_analyzed_template_btn
        gr.update(interactive=not locked),  # refresh_templates_btn
        gr.update(interactive=not locked),  # apply_template_btn
        gr.update(interactive=not locked),  # delete_template_btn
        gr.update(interactive=not locked),  # ref_attach_btn
        gr.update(interactive=not locked),  # ref_load_btn
        gr.update(interactive=not locked),  # templates_dropdown
        gr.update(interactive=not locked),  # ref_template_dropdown
        gr.update(interactive=not locked),  # up_file
        (reason if locked else ""),         # load_template_status
        gr.update(value=None, visible=not locked)  # generated_template_file
    )


def supa_plan_badge(sess):
    """
    Returns a small markdown badge for the user's plan.
    RULE: If profiles.is_paused is true -> ALWAYS show ⏸️ Paused (highest priority).
    Only when un-paused do we show Pro/Trial/Free.
    """
    if not sess:
        return "🆓 Free"

    try:
        # Make sure a profile row exists (safe, idempotent)
        ensure_profile_row(sess)

        # Use your existing profile loader (no extra authed helper needed)
        prow, _msg = supa_get_profile(sess)
        prow = prow or {}

        # 1) PAUSED OVERRIDES EVERYTHING
        if bool(prow.get("is_paused", False)) is True:
            return "⏸️ Paused"

        # 2) Otherwise: plan / subscription badge logic (best-effort)
        plan = (prow.get("plan") or prow.get("subscription_plan") or "").strip().lower()
        is_pro = bool(prow.get("is_pro", False))

        if plan in ("pro", "paid", "premium") or is_pro:
            return "⭐ Pro"
        if plan in ("trial", "free_trial"):
            return "🧪 Trial"
        return "🆓 Free"

    except Exception:
        # Never break the UI because of a badge
        return "🆓 Free"


def supa_save_profile(sess, display_name: str, notify_weekly: bool, notify_export_done: bool, notify_low_credits: bool):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return f"❌ {err}"

    try:
        supabase.postgrest.auth(access_token)
        now = datetime.now(timezone.utc).isoformat()

        payload = {
            "user_id": user_id,
            "display_name": (display_name or "").strip() or None,
            "notify_weekly_summary": bool(notify_weekly),
            "notify_export_done": bool(notify_export_done),
            "notify_low_credits": bool(notify_low_credits),
            "updated_at": now,
        }

        supabase.table("profiles").upsert(payload, on_conflict="user_id").execute()
        return "✅ Profile saved."
    except Exception as e:
        return f"❌ Failed to save profile: {type(e).__name__}: {e}"


# =============================
# AVATAR CROPPING (server-side)
# =============================
def _crop_avatar_square(in_path: str, out_path: str, size: int = 512, y_frac: float = 0.35):
    """
    Center-square crop, then resize to size x size.
    y_frac shifts crop window vertically:
      0.0 = bias up, 0.5 = centered, 1.0 = bias down
    """
    from PIL import Image

    im = Image.open(in_path).convert("RGB")
    w, h = im.size
    side = min(w, h)

    # Horizontal crop centered
    left = (w - side) // 2

    # Vertical crop with bias
    max_top = h - side
    top = int(max_top * max(0.0, min(1.0, y_frac)))

    box = (left, top, left + side, top + side)
    im2 = im.crop(box).resize((size, size), Image.LANCZOS)

    im2.save(out_path, format="PNG", optimize=True)
    return out_path


def _crop_avatar_face(in_path: str, out_path: str, size: int = 512, margin: float = 0.55):
    """
    Face-detect crop (largest face), expand by margin, then square-crop and resize.
    Returns out_path if face found, else None.
    """
    try:
        import cv2
        from PIL import Image
        import numpy as np

        img = cv2.imread(in_path)
        if img is None:
            return None

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        face_cascade = cv2.CascadeClassifier(cascade_path)
        faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(60, 60))

        if faces is None or len(faces) == 0:
            return None

        # pick largest face
        x, y, w, h = max(faces, key=lambda f: f[2] * f[3])

        # expand box
        cx = x + w / 2.0
        cy = y + h / 2.0
        side = max(w, h) * (1.0 + margin)

        x1 = int(cx - side / 2.0)
        y1 = int(cy - side / 2.0)
        x2 = int(cx + side / 2.0)
        y2 = int(cy + side / 2.0)

        H, W = img.shape[:2]
        x1 = max(0, x1); y1 = max(0, y1)
        x2 = min(W, x2); y2 = min(H, y2)

        crop = img[y1:y2, x1:x2]
        if crop.size == 0:
            return None

        # Convert to PIL, force square by center-crop if needed
        pil = Image.fromarray(cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)).convert("RGB")
        cw, ch = pil.size
        s = min(cw, ch)
        left = (cw - s) // 2
        top = (ch - s) // 2
        pil = pil.crop((left, top, left + s, top + s)).resize((size, size), Image.LANCZOS)

        pil.save(out_path, format="PNG", optimize=True)
        return out_path
    except Exception:
        return None


# =============================
# DEFAULT AVATAR (ADMIN)
# =============================
DEFAULT_AVATAR_BUCKET = "avatars"
DEFAULT_AVATAR_PATH = "system/default-avatar.png"


def _signed_storage_url(bucket_name: str, storage_path: str, expires_in: int = 300):
    """
    Safe helper: create a signed URL for any file in Supabase Storage.
    Returns None if signing fails or storage_path is blank.
    """
    storage_path = (storage_path or "").strip()
    if not storage_path:
        return None

    try:
        bucket = supabase.storage.from_(bucket_name)
        signed = bucket.create_signed_url(storage_path, expires_in)

        signed_url = (
            signed.get("signedURL")
            or signed.get("signedUrl")
            or signed.get("signed_url")
            or ""
        ).strip()

        return signed_url or None
    except Exception:
        return None


def supa_set_default_avatar(sess, avatar_file):
    """
    Admin-only:
    Uploads/replaces the system default avatar in Supabase Storage.

    Target:
      avatars / system/default-avatar.png

    Returns:
      (status_text, signed_url_or_none)
    """
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return f"❌ {err}", None

    if not _is_admin_user(sess):
        return "❌ Admin access required.", None

    if not avatar_file:
        return "❌ Please upload an image (PNG/JPG/WEBP).", None

    # Gradio File can be dict-like or path-like depending on version
    if isinstance(avatar_file, dict):
        fpath = avatar_file.get("name") or avatar_file.get("path") or ""
        fname = os.path.basename(fpath) if fpath else "default-avatar.png"
    else:
        fpath = str(avatar_file)
        fname = os.path.basename(fpath) if fpath else "default-avatar.png"

    if not fpath or not os.path.exists(fpath):
        return "❌ Uploaded default avatar file path not found.", None

    ext = os.path.splitext(fname)[1].lower()
    if ext not in {".png", ".jpg", ".jpeg", ".webp"}:
        return "❌ Supported formats: PNG, JPG, WEBP.", None

    try:
        import tempfile

        # IMPORTANT: use the session-authenticated client, not the global client
        sb = _sb_authed_from_session(sess)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.close()

        # Try face crop first; fallback to square crop
        outp = _crop_avatar_face(fpath, tmp.name, size=512, margin=0.65)
        if not outp:
            _crop_avatar_square(fpath, tmp.name, size=512, y_frac=0.35)

        with open(tmp.name, "rb") as f:
            data = f.read()

        bucket = sb.storage.from_(DEFAULT_AVATAR_BUCKET)
        bucket.upload(
            DEFAULT_AVATAR_PATH,
            data,
            {"content-type": "image/png", "x-upsert": "true"}
        )

        signed = bucket.create_signed_url(DEFAULT_AVATAR_PATH, 300)
        signed_url = (
            signed.get("signedURL")
            or signed.get("signedUrl")
            or signed.get("signed_url")
            or ""
        ).strip()

        return "✅ Default avatar saved.", (signed_url or None)

    except Exception as e:
        return f"❌ Failed to save default avatar: {type(e).__name__}: {e}", None


def supa_set_avatar(sess, avatar_file):
    """
    Uploads an avatar image to private Storage bucket 'avatars',
    stores the storage path on public.profiles.avatar_storage_path,
    and returns (status_text, signed_url_or_none) for preview.

    Returns:
      (status, avatar_preview_url_or_none)
    """
    # Ensure profile row exists
    ensure_profile_row(sess)

    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return f"❌ {err}", None

    if not avatar_file:
        return "❌ Please upload an image (PNG/JPG/WEBP).", None

    # Gradio File can be dict-like or path-like depending on version
    if isinstance(avatar_file, dict):
        fpath = avatar_file.get("name") or avatar_file.get("path") or ""
        fname = os.path.basename(fpath) if fpath else "avatar.png"
    else:
        fpath = str(avatar_file)
        fname = os.path.basename(fpath) if fpath else "avatar.png"

    if not fpath or not os.path.exists(fpath):
        return "❌ Uploaded avatar file path not found.", None

    ext = os.path.splitext(fname)[1].lower()
    if ext not in {".png", ".jpg", ".jpeg", ".webp"}:
        return "❌ Supported formats: PNG, JPG, WEBP.", None

    try:
        supabase.postgrest.auth(access_token)

        # --- Crop/Zoom avatar server-side (face if possible, else safe square) ---
        import tempfile

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.close()

        # 1) Try face crop
        outp = _crop_avatar_face(fpath, tmp.name, size=512, margin=0.65)

        # 2) Fallback to square crop if no face found
        if not outp:
            _crop_avatar_square(fpath, tmp.name, size=512, y_frac=0.35)

        # Read cropped PNG bytes
        with open(tmp.name, "rb") as f:
            data = f.read()

        # Always store as PNG (stable + consistent)
        storage_path = f"{user_id}/avatar.png"
        ctype = "image/png"

        bucket = supabase.storage.from_("avatars")

        bucket.upload(
            storage_path,
            data,
            {"content-type": ctype, "x-upsert": "true"}
        )

        # Save path on profile
        supabase.table("profiles").update({
            "avatar_storage_path": storage_path,
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }).eq("user_id", user_id).execute()

        # Signed URL for preview (5 min)
        signed = bucket.create_signed_url(storage_path, 300)
        signed_url = (
            signed.get("signedURL")
            or signed.get("signedUrl")
            or signed.get("signed_url")
            or ""
        ).strip()

        return "✅ Avatar saved.", (signed_url or None)

    except Exception as e:
        return f"❌ Failed to save avatar: {type(e).__name__}: {e}", None


def supa_load_avatar(sess):
    """
    Loads the user's avatar from profiles.avatar_storage_path.
    If the user has not uploaded one yet, falls back to the system default avatar.

    Returns:
      (status, avatar_preview_url_or_none)
    """
    ensure_profile_row(sess)

    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return f"❌ {err}", None

    try:
        supabase.postgrest.auth(access_token)

        res = (
            supabase.table("profiles")
            .select("avatar_storage_path")
            .eq("user_id", user_id)
            .limit(1)
            .execute()
        )
        rows = getattr(res, "data", None) or []
        if not rows:
            return "ℹ️ No profile row found.", None

        storage_path = (rows[0].get("avatar_storage_path") or "").strip()

        # 1) Personal avatar if present
        if storage_path:
            signed_url = _signed_storage_url("avatars", storage_path, expires_in=300)
            if signed_url:
                return "✅ Avatar loaded.", signed_url

        # 2) Fallback to system default avatar
        default_url = _signed_storage_url(DEFAULT_AVATAR_BUCKET, DEFAULT_AVATAR_PATH, expires_in=300)
        if default_url:
            return "✅ Default avatar loaded.", default_url

        return "ℹ️ No avatar uploaded yet, and no default avatar is set.", None

    except Exception as e:
        return f"❌ Failed to load avatar: {type(e).__name__}: {e}", None


def _email_from_session(sess) -> str:
    """
    Best-effort extraction of email from supabase_session state.
    Works across common shapes (dict with user/email).
    """
    try:
        if not sess:
            return ""
        # common shapes we've seen in apps: sess["user"]["email"] or sess["email"]
        if isinstance(sess, dict):
            u = sess.get("user") or {}
            if isinstance(u, dict) and u.get("email"):
                return (u.get("email") or "").strip()
            if sess.get("email"):
                return (sess.get("email") or "").strip()
        return ""
    except Exception:
        return ""


def _nice_default_name_from_email(email: str) -> str:
    """
    "jane.doe@youremail.com" -> "Jane Doe"
    "jane" -> "Jane"
    """
    e = (email or "").strip()
    if not e or "@" not in e:
        return ""
    local = e.split("@", 1)[0]
    local = local.replace(".", " ").replace("_", " ").replace("-", " ").strip()
    if not local:
        return ""
    # Title-case each word
    return " ".join([w[:1].upper() + w[1:] for w in local.split() if w])




def supa_save_profile_v2(
    sess,

    # New Account Details fields
    first_name,
    last_name,
    email_ui,
    country,
    language,
    years_taught,
    school_type,
    curriculum_system,
    subjects_taught,
    assessment_style,
    difficulty_pref,
    spelling_pref,
    marking_style,
    preferred_tone,
    class_size,
    ability_mix,
    confirm_educator,
    confirm_review,

    # Existing fields (still used)
    display_name,
    notify_weekly,
    notify_export_done,
    notify_low_credits
):
    """
    V2 profile saver.
    Must accept the full UI input list (even if some are UI-only today),
    so the app never crashes.

    Writes:
    - display_name + notify flags (existing behavior)
    - additional fields into profile row (best-effort; if DB columns missing, it fails gracefully)
    """

    # Ensure row exists (idempotent)
    ensure_profile_row(sess)

    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return f"❌ {err}"

    try:
        supabase.postgrest.auth(access_token)

        # Build payload (only include keys that are safe)
        payload = {
            # Existing fields
            "display_name": (display_name or "").strip(),
            "notify_weekly_summary": bool(notify_weekly),
            "notify_export_done": bool(notify_export_done),
            "notify_low_credits": bool(notify_low_credits),

            # New Account Details fields (best-effort)
            "first_name": (first_name or "").strip(),
            "last_name": (last_name or "").strip(),
            "country": (country or "").strip(),
            "language": (language or "").strip(),
            "years_taught": years_taught or [],
            "school_type": (school_type or "").strip(),
            "curriculum_system": (curriculum_system or "").strip(),
            "subjects_taught": subjects_taught or [],
            "assessment_style": assessment_style or [],
            "difficulty_pref": (difficulty_pref or "").strip(),
            "spelling_pref": (spelling_pref or "").strip(),
            "marking_style": (marking_style or "").strip(),
            "preferred_tone": (preferred_tone or "").strip(),
            "class_size": (class_size or "").strip(),
            "ability_mix": (ability_mix or "").strip(),
            "confirm_educator": bool(confirm_educator),
            "confirm_review": bool(confirm_review),
        }

        # DO NOT trust email_ui textbox. Real email is derived from session.
        real_email = _email_from_session(sess) or ""
        if real_email:
            payload["email"] = real_email

        # Attempt update
        # If your DB does NOT have some of these columns yet, Supabase may error.
        # We'll retry with a reduced payload (only the known-good legacy fields).
        try:
            supabase.table("profiles").update(payload).eq("user_id", user_id).execute()
        except Exception as e_cols:
            legacy = {
                "display_name": payload["display_name"],
                "notify_weekly_summary": payload["notify_weekly_summary"],
                "notify_export_done": payload["notify_export_done"],
                "notify_low_credits": payload["notify_low_credits"],
            }
            supabase.table("profiles").update(legacy).eq("user_id", user_id).execute()
            return "⚠️ Saved core profile settings. (Extra Account Details fields not saved yet — DB columns missing.)"

        return "✅ Profile settings saved."

    except Exception as e:
        return f"❌ Save failed: {type(e).__name__}: {e}"


def supa_profile_banner(sess):
    """
    Returns (banner_avatar_url, banner_name_md, banner_signed_md)

    Priority:
    1) User's personal avatar from profiles.avatar_storage_path
    2) System default avatar from Supabase Storage
    """
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return (None, "**Not signed in**", "_Not signed in._")

    try:
        supabase.postgrest.auth(access_token)

        email = _email_from_session(sess) or ""
        signed_md = f"_Signed in as: `{email}`_" if email else f"_Signed in (user_id: {user_id})_"

        prow, _ = supa_get_profile(sess)
        prow = prow or {}

        # Preferred banner name order:
        # 1) first_name + last_name
        # 2) display_name
        # 3) nice name from email
        first_name = (prow.get("first_name") or "").strip()
        last_name = (prow.get("last_name") or "").strip()
        display_name = (prow.get("display_name") or "").strip()

        full_name = " ".join([p for p in [first_name, last_name] if p]).strip()

        banner_name = full_name or display_name
        if not banner_name and email:
            banner_name = _nice_default_name_from_email(email) or ""

        name_md = f"**{banner_name or '—'}**"

        # 1) Personal avatar first
        storage_path = (prow.get("avatar_storage_path") or "").strip()
        if storage_path:
            signed_url = _signed_storage_url("avatars", storage_path, expires_in=300)
            if signed_url:
                return (signed_url, name_md, signed_md)

        # 2) Fallback to system default avatar
        default_url = _signed_storage_url(DEFAULT_AVATAR_BUCKET, DEFAULT_AVATAR_PATH, expires_in=300)
        if default_url:
            return (default_url, name_md, signed_md)

        # 3) Nothing available
        return (None, name_md, signed_md)

    except Exception as e:
        return (None, "**Profile load failed**", f"_Error: {type(e).__name__}: {e}_")


def supa_global_banner_payload(sess):
    """
    Returns (banner_visible, avatar_url, name_md, signed_md)
    """
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return (False, None, "**Not signed in**", "_Not signed in._")

    avatar_url, name_md, signed_md = supa_profile_banner(sess)
    return (True, avatar_url, name_md, signed_md)


def supa_profile_snapshot(sess):
    """
    Returns (MUST match Gradio outputs order):

    (
      status,
      avatar_url,
      header,

      email,
      first_name,
      last_name,
      country,
      language,
      years_taught,
      school_type,
      curriculum_system,
      subjects_taught,
      assessment_style,
      difficulty_pref,
      spelling_pref,
      marking_style,
      preferred_tone,
      class_size,
      ability_mix,
      confirm_educator,
      confirm_review,

      display_name,
      notify_weekly,
      notify_export_done,
      notify_low_credits,

      drafts_count_text,
      drafts_list_text,
      templates_count_text,
      templates_list_text,

      privacy_ack_status,
      terms_ack_status,
      tcs_ack_status,

      privacy_ack_cb_update,
      terms_ack_cb_update,
      tcs_ack_cb_update
    )
    """
    ensure_profile_row(sess)

    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return (
            f"❌ {err}",
            None,
            "Not signed in.",

            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", False, False,

            "", False, True, True,
            "Drafts: ?", "",
            "Templates: ?", "",

            # Legal
            "❌ Not acknowledged yet",
            "❌ Not acknowledged yet",
            "❌ Not acknowledged yet",

            gr.update(visible=True, value=False),
            gr.update(visible=True, value=False),
            gr.update(visible=True, value=False),
            gr.update(visible=False),
        )

    try:
        supabase.postgrest.auth(access_token)

        email = _email_from_session(sess) or ""
        header = f"Signed in as: `{email}`" if email else f"Signed in (user_id: {user_id})"

        prow, _ = supa_get_profile(sess)
        prow = prow or {}

        ## Avatar (signed with fallback)
        avatar_url = None
        storage_path = (prow.get("avatar_storage_path") or "").strip()

        # 1) Personal avatar first
        if storage_path:
            avatar_url = _signed_storage_url("avatars", storage_path, expires_in=300)

        # 2) Fallback to system default avatar
        if not avatar_url:
            avatar_url = _signed_storage_url(DEFAULT_AVATAR_BUCKET, DEFAULT_AVATAR_PATH, expires_in=300)

        # --- Account details ---
        first_name = prow.get("first_name", "") or ""
        last_name = prow.get("last_name", "") or ""
        country = prow.get("country", "") or ""
        language = prow.get("language", "") or ""
        years_taught = prow.get("years_taught", "") or ""
        school_type = prow.get("school_type", "") or ""
        curriculum_system = prow.get("curriculum_system", "") or ""
        subjects_taught = prow.get("subjects_taught", "") or ""
        assessment_style = prow.get("assessment_style", "") or ""
        difficulty_pref = prow.get("difficulty_pref", "") or ""
        spelling_pref = prow.get("spelling_pref", "") or ""
        marking_style = prow.get("marking_style", "") or ""
        preferred_tone = prow.get("preferred_tone", "") or ""
        class_size = prow.get("class_size", "") or ""
        ability_mix = prow.get("ability_mix", "") or ""
        confirm_educator = bool(prow.get("confirm_educator", False))
        confirm_review = bool(prow.get("confirm_review", False))

        # Existing
        display_name = prow.get("display_name", "") or ""
        if not display_name and email:
            display_name = _nice_default_name_from_email(email) or ""

        notify_weekly = bool(prow.get("notify_weekly_summary", False))
        notify_export_done = bool(prow.get("notify_export_done", True))
        notify_low_credits = bool(prow.get("notify_low_credits", True))

        # Drafts
        dres = supabase.table("drafts").select("id", count="exact").eq("user_id", user_id).execute()
        drafts_count_text = f"Drafts: {getattr(dres, 'count', 0) or 0}"

        dlist = (
            supabase.table("drafts")
            .select("title,subject,output_type,updated_at")
            .eq("user_id", user_id)
            .order("updated_at", desc=True)
            .limit(10)
            .execute()
        )
        # Group drafts by output_type (for "Your Saved Work")
        rows = (dlist.data or [])

        grouped = {}
        for r in rows:
            ot = (r.get("output_type") or "Uncategorised").strip() or "Uncategorised"
            title_txt = (r.get("title") or "Untitled").strip()
            subj_txt = (r.get("subject") or "").strip()

            line = f"• {title_txt}" + (f" — {subj_txt}" if subj_txt else "")
            grouped.setdefault(ot, []).append(line)

        # Preferred display order (matches your Workspace dropdown)
        preferred_order = [
            "Test / Quiz",
            "Worksheet",
            "Investigation",
            "Lesson",
            "PowerPoint lesson",
            "Marking key / Memo only",
            "Rubric",
            "Homework",
            "Exam",
            "Lesson Plan",
            "Revision Sheet",
            "Custom",
        ]

        sections = []
        for ot in preferred_order:
            if ot in grouped:
                sections.append(f"### {ot}\n" + "\n".join(grouped[ot]))

        # Any other unexpected types at the bottom (sorted)
        extras = sorted([k for k in grouped.keys() if k not in preferred_order])
        for ot in extras:
            sections.append(f"### {ot}\n" + "\n".join(grouped[ot]))

        drafts_list_text = "\n\n".join(sections) if sections else "(No drafts yet)"

        # Templates
        tres = supabase.table("templates").select("id", count="exact").eq("user_id", user_id).execute()
        templates_count_text = f"Templates: {getattr(tres, 'count', 0) or 0}"

        tlist = (
            supabase.table("templates")
            .select("name,subject,category,updated_at,created_at")
            .eq("user_id", user_id)
            .order("updated_at", desc=True)
            .limit(10)
            .execute()
        )

        trows = (tlist.data or [])

        # Group templates by Subject -> Category -> Names
        tgrouped = {}  # {subject: {category: [name, ...]}}
        for r in trows:
            subj = (r.get("subject") or "Uncategorised").strip() or "Uncategorised"
            cat  = (r.get("category") or "Custom").strip() or "Custom"
            name = (r.get("name") or "Untitled Template").strip()

            if subj not in tgrouped:
                tgrouped[subj] = {}
            tgrouped[subj].setdefault(cat, []).append(name)


        # Canonical subject order (match Workspace Subject dropdown)
        canonical_subject_order = [
            "English",
            "Mathematics",
            "Mathematics (Methods)",
            "Mathematics (Applications)",
            "Mathematics (General)",
            "Science",
            "Biology",
            "Chemistry",
            "Physics",
            "Human Biology",
            "HASS / Humanities",
            "Geography",
            "History",
            "Economics",
            "Accounting & Finance",
            "Business / Commerce",
            "Digital Technologies / Computing",
            "Design & Technology",
            "Health & Physical Education",
            "Drama",
            "Music",
            "Visual Arts",
            "Languages",
        ]

        tsections = []

        # Category order (match TEMPLATE_CATEGORIES first, then anything unexpected)
        preferred_category_order = list(TEMPLATE_CATEGORIES)

        def _render_subject_section(subj: str, cats: dict) -> str:
            # cats = {category: [name, name, ...]}
            total = sum(len(v) for v in (cats or {}).values())

            # Subject header line stays the same (your UI relies on this style)
            lines = [f"### {subj} ({total})"]

            # Build a stable category order: preferred first, then unexpected ones
            preferred = list(preferred_category_order)
            extra_cats = sorted(
                [c for c in (cats or {}).keys() if c not in preferred],
                key=lambda s: s.lower()
            )
            ordered_cats = preferred + extra_cats

            body_lines = []
            for cat in ordered_cats:
                names = (cats or {}).get(cat) or []
                if not names:
                    continue

                # Category “heading” line (stands out)
                body_lines.append(f"**▸ {cat} ({len(names)})**")

                # Template names indented under the category
                for n in names:
                    body_lines.append(f"        • {n}")

                body_lines.append("")  # blank line between categories

            if not body_lines:
                body_lines = ["(None yet)"]

            lines.append("\n".join(body_lines).rstrip())
            return "\n".join(lines)

        # 1) First: canonical order (only if present)
        for subj in canonical_subject_order:
            if subj in tgrouped:
                tsections.append(_render_subject_section(subj, tgrouped[subj]))

        # 2) Then: any extra / custom subjects (alphabetical), excluding "Uncategorised"
        extras = sorted(
            [k for k in tgrouped.keys() if k not in canonical_subject_order and k != "Uncategorised"],
            key=lambda s: s.lower()
        )
        for subj in extras:
            tsections.append(_render_subject_section(subj, tgrouped[subj]))

        # 3) Finally: Uncategorised at the bottom (if present)
        if "Uncategorised" in tgrouped:
            tsections.append(_render_subject_section("Uncategorised", tgrouped["Uncategorised"]))

        templates_list_text = "\n\n".join(tsections) if tsections else "(No templates yet)"

        # =========================
        # LEGAL ACKS (Privacy / Terms / T&Cs)
        # =========================

        def _fmt_ack(dt: str, ver: str) -> str:
            dt = (dt or "").strip()
            ver = (ver or "").strip()
            if dt:
                # Keep it simple + readable (don’t risk parsing formats yet)
                if ver:
                    return f"✅ Acknowledged on **{dt}** (v{ver})"
                return f"✅ Acknowledged on **{dt}**"
            return "❌ Not acknowledged yet"

        pp_at  = (prow.get("privacy_policy_ack_at") or "").strip()
        pp_ver = (prow.get("privacy_policy_version") or "").strip()

        tou_at  = (prow.get("terms_of_use_ack_at") or "").strip()
        tou_ver = (prow.get("terms_of_use_version") or "").strip()

        tcs_at  = (prow.get("terms_and_conditions_ack_at") or "").strip()
        tcs_ver = (prow.get("terms_and_conditions_version") or "").strip()

        # Required versions for current legal docs
        # Source of truth = Supabase table: public.legal_config
        # Safety fallback = LEGAL_VERSIONS (code) -> then hardcoded defaults
        pp_req  = None
        tou_req = None
        tcs_req = None

        try:
            # Fetch required versions from DB (expects 3 rows: privacy/terms/tcs)
            cfg_rows = (
                supabase.table("legal_config")
                .select("doc_type,current_version")
                .in_("doc_type", ["privacy", "terms", "tcs"])
                .execute()
            )
            cfg = {r.get("doc_type"): r.get("current_version") for r in (cfg_rows.data or [])}

            pp_req  = (cfg.get("privacy") or "").strip() or None
            tou_req = (cfg.get("terms") or "").strip() or None
            tcs_req = (cfg.get("tcs") or "").strip() or None
        except Exception:
            # Ignore DB errors; fall back below
            pass

        # Fallback chain
        if not pp_req:
            pp_req = (LEGAL_VERSIONS.get("privacy") if isinstance(globals().get("LEGAL_VERSIONS"), dict) else None) or "PP-2026-02-18"
        if not tou_req:
            tou_req = (LEGAL_VERSIONS.get("terms") if isinstance(globals().get("LEGAL_VERSIONS"), dict) else None) or "TOU-2026-02-18"
        if not tcs_req:
            tcs_req = (LEGAL_VERSIONS.get("tcs") if isinstance(globals().get("LEGAL_VERSIONS"), dict) else None) or "TCS-2026-02-18"

        # Only "acknowledged" if user acknowledged the CURRENT required version
        pp_ok  = bool(pp_at) and (pp_ver == pp_req)
        tou_ok = bool(tou_at) and (tou_ver == tou_req)
        tcs_ok = bool(tcs_at) and (tcs_ver == tcs_req)

        pp_status  = _fmt_ack(pp_at, pp_ver) if pp_ok else "❌ Not acknowledged yet"
        tou_status = _fmt_ack(tou_at, tou_ver) if tou_ok else "❌ Not acknowledged yet"
        tcs_status = _fmt_ack(tcs_at, tcs_ver) if tcs_ok else "❌ Not acknowledged yet"

        # Profile checkboxes:
        # - hide if already acknowledged for the current version
        # - show only if still outstanding
        pp_cb_update = gr.update(visible=not pp_ok, value=False)
        tou_cb_update = gr.update(visible=not tou_ok, value=False)
        tcs_cb_update = gr.update(visible=not tcs_ok, value=False)
        legal_up_to_date_note_update = gr.update(visible=(pp_ok and tou_ok and tcs_ok))

        return (
            "✅ Profile loaded.",
            avatar_url,
            header,

            email,
            first_name,
            last_name,
            country,
            language,
            years_taught,
            school_type,
            curriculum_system,
            subjects_taught,
            assessment_style,
            difficulty_pref,
            spelling_pref,
            marking_style,
            preferred_tone,
            class_size,
            ability_mix,
            confirm_educator,
            confirm_review,

            display_name,
            notify_weekly,
            notify_export_done,
            notify_low_credits,

            drafts_count_text,
            drafts_list_text,
            templates_count_text,
            templates_list_text,

            # Legal
            pp_status,
            tou_status,
            tcs_status,

            pp_cb_update,
            tou_cb_update,
            tcs_cb_update,
            legal_up_to_date_note_update,
        )

    except Exception as e:
        return (
            f"❌ Failed to load snapshot: {type(e).__name__}: {e}",
            None,
            "Profile load failed.",

            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", False, False,

            "", False, True, True,
            "Drafts: ?", "",
            "Templates: ?", "",

            # Legal
            "❌ Not acknowledged yet",
            "❌ Not acknowledged yet",
            "❌ Not acknowledged yet",
        )

def supa_profile_summary_md(sess):
    ok, _ = ensure_profile_row(sess)
    if not ok:
        return "**Saved profile summary**\n\n_Not available._"

    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return "**Saved profile summary**\n\n_Not signed in._"

    try:
        supabase.postgrest.auth(access_token)
        res = supabase.table("profiles").select("*").eq("user_id", user_id).limit(1).execute()
        row = (res.data or [{}])[0] if res.data else {}

        def show(v, fallback="—"):
            if v is None:
                return fallback
            if isinstance(v, str) and not v.strip():
                return fallback
            if isinstance(v, list):
                return ", ".join(str(x) for x in v) if v else fallback
            if isinstance(v, bool):
                return "Yes" if v else "No"
            return str(v)

        return f"""**Saved profile summary**

**First name:** {show(row.get("first_name"))}  
**Last name:** {show(row.get("last_name"))}  
**Email:** {show(row.get("email"))}  
**Country:** {show(row.get("country"))}  
**Language:** {show(row.get("language"))}  
**Years taught:** {show(row.get("years_taught"))}  
**School type:** {show(row.get("school_type"))}  
**Curriculum system:** {show(row.get("curriculum_system"))}  
**Subjects taught:** {show(row.get("subjects_taught"))}  
**Assessment style:** {show(row.get("assessment_style") or row.get("assessment_styles"))}  
**Difficulty preference:** {show(row.get("difficulty_pref"))}  
**Spelling preference:** {show(row.get("spelling_pref"))}  
**Marking style:** {show(row.get("marking_style"))}  
**Preferred tone:** {show(row.get("preferred_tone") or row.get("tone_pref"))}  
**Class size:** {show(row.get("class_size"))}  
**Ability mix:** {show(row.get("ability_mix"))}  
**Educator confirmed:** {show(row.get("confirm_educator"))}  
**Review confirmed:** {show(row.get("confirm_review"))}
"""
    except Exception as e:
        return f"**Saved profile summary**\n\n❌ Failed: {type(e).__name__}: {e}"


def _file_ext(path: str) -> str:
    return os.path.splitext(path or "")[1].lower().strip()


def _read_text_file(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read().strip()


# ===============================
# PROJECT ROOT (GLOBAL)
# ===============================
PROJECT_ROOT = os.getenv("PROJECT_ROOT", "").strip()
if not PROJECT_ROOT:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))


# ===============================
# LEGAL DOCUMENT LOADING (GLOBAL)
# ===============================
LEGAL_MD_PATHS = {
    "privacy": os.path.join(PROJECT_ROOT, "legal", "privacy_policy.md"),
    "terms": os.path.join(PROJECT_ROOT, "legal", "terms_of_use.md"),
    "tcs": os.path.join(PROJECT_ROOT, "legal", "terms_and_conditions.md"),
}

LEGAL_MD_TEXT = {
    k: _read_text_file(p) for k, p in LEGAL_MD_PATHS.items()
}


def _docx_to_markdown_via_pandoc(docx_path: str) -> str:
    tmpdir = tempfile.mkdtemp()
    out_md = os.path.join(tmpdir, "template.md")
    subprocess.run(["pandoc", docx_path, "-t", "gfm", "-o", out_md], check=True)
    return _read_text_file(out_md)


def upload_template_scaffold(
    session_state,
    up_template_name: str,
    up_template_desc: str,
    up_template_category: str,
    up_file,
    up_is_public: bool = False
):
    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return f"❌ {err}", gr.update(), f"❌ {err}"

    name = (up_template_name or "").strip()
    if not name:
        return "❌ Template name is required.", gr.update(), "❌ Template name is required."

    if not up_file:
        return "❌ Please upload a .md, .txt, or .docx file.", gr.update(), "❌ No file uploaded."

    if isinstance(up_file, dict):
        file_path = up_file.get("name") or up_file.get("path") or ""
    else:
        file_path = str(up_file)

    if not file_path or not os.path.exists(file_path):
        return "❌ Uploaded file path not found.", gr.update(), "❌ File path not found."

    ext = _file_ext(file_path)
    if ext not in ALLOWED_SCAFFOLD_EXTS:
        return (
            "❌ Unsupported file type for scaffold. Upload .md, .txt, or .docx.",
            gr.update(),
            "❌ Unsupported file type."
        )

    try:
        if ext in {".md", ".txt"}:
            scaffold_md = _read_text_file(file_path)
        else:
            scaffold_md = _docx_to_markdown_via_pandoc(file_path)

        if not scaffold_md.strip():
            return "❌ The uploaded file converted to empty content.", gr.update(), "❌ Empty content."

        template_data = {
            "reference_storage_path": None,
            "reference_file_url": None,
            "reference_file_name": None,
            "reference_file_type": None,
            "user_id": user_id,
            "name": name,
            "description": (up_template_desc or "").strip(),

            # ✅ Upload scaffolds usually don’t have a subject (until we add a UI for it later)
            "subject": None,

            # ✅ Category comes from that dropdown
            "category": (up_template_category or "Custom").strip() or "Custom",

            "template_md": scaffold_md,
            "template_ppt": "",
            "is_public": bool(up_is_public),
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }

        supabase.table("templates").insert(template_data).execute()

        choices, msg = list_my_templates(session_state, category_filter="All")
        dd_update = gr.update(choices=choices, value=(choices[0] if choices else None))

        return f"✅ Uploaded & saved template scaffold: {name}", dd_update, msg

    except Exception as e:
        if "duplicate" in str(e).lower():
            return f"❌ Template name '{name}' already exists. Choose a different name.", gr.update(), "❌ Duplicate name."
        return f"❌ Upload failed: {type(e).__name__}: {e}", gr.update(), "❌ Upload failed."


def save_as_template(
    session_state,
    template_name: str,
    template_desc: str,
    template_category: str,
    draft_subject: str,              # ✅ NEW
    markdown_content: str,
    ppt_content: str,
    is_public: bool = False
) -> str:
    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return f"❌ {err}"

    if not template_name.strip():
        return "❌ Template name is required."

    if not markdown_content.strip():
        return "❌ No content to save as template."

    try:
        template_data = {
            "reference_storage_path": None,
            "reference_file_url": None,
            "reference_file_name": None,
            "reference_file_type": None,
            "user_id": user_id,
            "name": template_name.strip(),
            "description": (template_desc or "").strip(),

            # ✅ Subject = the real school subject from Workspace (draft_subject_state)
            "subject": (draft_subject or "").strip() or None,

            # ✅ Category = Worksheet / Lesson Plan / Rubric etc
            "category": (template_category or "Custom").strip() or "Custom",

            "template_md": markdown_content,
            "template_ppt": ppt_content or "",
            "is_public": is_public,
            "updated_at": datetime.now(timezone.utc).isoformat()
        }

        supabase.table("templates").insert(template_data).execute()
        return f"✅ Template '{template_name}' saved successfully!"

    except Exception as e:
        if "duplicate key" in str(e).lower():
            return f"❌ Template name '{template_name}' already exists. Choose a different name."
        return f"❌ Failed to save template: {type(e).__name__}: {e}"


def list_my_templates(session_state, category_filter: str = ""):
    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return [], f"❌ {err}"

    try:
        query = supabase.table("templates") \
            .select("id,name,description,subject,created_at")\
            .eq("user_id", user_id)

        if category_filter and category_filter != "All":
            query = query.eq("category", category_filter)

        query = query.order("created_at", desc=True)
        res = query.execute()
        rows = getattr(res, "data", None) or []

        choices = []
        for r in rows:
            template_id = r.get("id", "")
            name = r.get("name", "Unnamed Template")
            choices.append(f"{name} | {template_id}")

        return choices, f"✅ Found {len(choices)} templates."

    except Exception as e:
        return [], f"❌ Failed to load templates: {type(e).__name__}: {e}"


def load_template(session_state, template_choice):
    if isinstance(template_choice, list):
        template_choice = template_choice[0] if template_choice else ""

    if not template_choice:
        return "", "", "❌ No template selected."

    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return "", "", f"❌ {err}"

    try:
        if "|" in template_choice:
            parts = template_choice.split("|", 1)
            template_name = parts[0].strip()
            template_id = parts[1].strip()
        else:
            template_name = template_choice.strip()
            template_id = ""

        if template_id and len(template_id) >= 36:
            res = supabase.table("templates") \
                .select("template_md,template_ppt,name") \
                .eq("id", template_id) \
                .eq("user_id", user_id) \
                .limit(1) \
                .execute()
        else:
            res = supabase.table("templates") \
                .select("template_md,template_ppt,name") \
                .eq("user_id", user_id) \
                .eq("name", template_name) \
                .limit(1) \
                .execute()

        rows = getattr(res, "data", None) or []
        if not rows:
            return "", "", f"❌ Template not found or access denied."

        row = rows[0]
        return row.get("template_md", ""), row.get("template_ppt", ""), f"✅ Loaded template: {row.get('name', '')}"

    except Exception as e:
        return "", "", f"❌ Failed to load template: {type(e).__name__}: {e}"


def delete_template(session_state, template_choice: str) -> str:
    if isinstance(template_choice, list):
        template_choice = template_choice[0] if template_choice else ""

    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return f"❌ {err}"

    if "|" in template_choice:
        template_id = template_choice.split("|", 1)[1].strip()
    else:
        template_id = template_choice.strip()

    if not template_id:
        return "❌ No template selected."

    try:
        supabase.table("templates") \
            .delete() \
            .eq("id", template_id) \
            .eq("user_id", user_id) \
            .execute()

        return f"✅ Template deleted."

    except Exception as e:
        return f"❌ Failed to delete template: {type(e).__name__}: {e}"


def attach_reference_pdf_to_template(session_state, template_choice, pdf_file):
    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return f"❌ {err}", None

    template_id = _parse_template_id(template_choice)
    if not template_id:
        return "❌ Please select a template first.", None

    if not pdf_file:
        return "❌ Please upload a PDF file.", None

    if isinstance(pdf_file, dict):
        pdf_path = pdf_file.get("name") or pdf_file.get("path") or ""
        pdf_name = os.path.basename(pdf_path) if pdf_path else "reference.pdf"
    else:
        pdf_path = str(pdf_file)
        pdf_name = os.path.basename(pdf_path)

    if not pdf_path or not os.path.exists(pdf_path):
        return "❌ Uploaded PDF path not found.", None

    ext = os.path.splitext(pdf_name)[1].lower()
    if ext != ".pdf":
        return "❌ Only PDF files are supported here.", None

    try:
        with open(pdf_path, "rb") as f:
            data = f.read()

        # Safer, more structured upload path (avoids collisions and keeps refs organised)
        safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", pdf_name).strip("_") or "reference.pdf"
        ref_id = uuid.uuid4().hex[:10]
        storage_path = f"{user_id}/templates/{template_id}/reference/{ref_id}_{safe_name}"
        bucket = supabase.storage.from_("template_refs")
        bucket.upload(storage_path, data, {"content-type": "application/pdf", "x-upsert": "true"})
        
        # Bucket is PRIVATE in V2.5 → use signed URLs, never public URLs
        signed = bucket.create_signed_url(storage_path, 300)  # 5 minutes
        signed_url = (
            signed.get("signedURL")
            or signed.get("signedUrl")
            or signed.get("signed_url")
            or ""
        )
        
        supabase.table("templates").update({
            # New field (V2.5)
            "reference_storage_path": storage_path,
        
            # Legacy fields (keep for backwards compatibility; stop relying on them)
            "reference_file_url": None,
            
            "reference_file_name": pdf_name,
            "reference_file_type": "application/pdf",
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }).eq("id", template_id).eq("user_id", user_id).execute()
        
        # Return a signed link so the user can download/open immediately
        return "✅ Reference PDF attached to template.", (signed_url or None)


    except Exception as e:
        return f"❌ Failed to attach reference PDF: {type(e).__name__}: {e}", None


def load_reference_pdf_for_template(session_state, template_choice):
    """
    V2.5: Bucket is private. Prefer reference_storage_path + signed URL.
    Backwards compatible: if storage_path missing, fall back to legacy reference_file_url.
    IMPORTANT: return order MUST match Gradio outputs:
      1) status textbox
      2) file/url output
    """
    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return f"❌ {err}", None

    template_id = _parse_template_id(template_choice)
    if not template_id:
        return "❌ Please select a template first.", None

    try:
        res = (
            supabase.table("templates")
            .select("reference_storage_path, reference_file_url, reference_file_name")
            .eq("id", template_id)
            .eq("user_id", user_id)
            .limit(1)
            .execute()
        )

        rows = getattr(res, "data", None) or []
        if not rows:
            return "❌ Template not found or access denied.", None

        row = rows[0]
        storage_path = (row.get("reference_storage_path") or "").strip()
        legacy_url = (row.get("reference_file_url") or "").strip()
        fname = row.get("reference_file_name") or "reference.pdf"

        # Prefer V2.5 signed URL from storage path
        if storage_path:
            bucket = supabase.storage.from_("template_refs")
            signed = bucket.create_signed_url(storage_path, 300)  # 5 minutes

            signed_url = (
                signed.get("signedURL")
                or signed.get("signedUrl")
                or signed.get("signed_url")
                or ""
            ).strip()

            if signed_url:
                return f"✅ Loaded: {fname}", signed_url

            return "❌ Reference PDF exists but signed URL could not be created.", None

        # Backwards compatibility (older templates)
        if legacy_url:
            return f"✅ Loaded (legacy): {fname}", legacy_url

        return "ℹ️ No reference PDF attached to this template yet.", None

    except Exception as e:
        return f"❌ Failed to load reference PDF: {type(e).__name__}: {e}", None


# =============================
# SUPABASE DRAFT LIBRARY (CRUD)
# =============================
def _parse_version(label: str) -> int:
    text = (label or "").strip().lower()

    # Accept both:
    #   "v3"
    #   "Version 3 — 01 Feb 2026"
    m = re.match(r"v(\d+)", text)
    if not m:
        m = re.match(r"version\s+(\d+)", text)

    return int(m.group(1)) if m else 0

def ensure_profile_row(sess):
    """
    Ensure the logged-in user has a row in public.profiles.
    Safe to call repeatedly (idempotent).
    """
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return False, err

    try:
        supabase.postgrest.auth(access_token)

        # Insert row if missing; do nothing if it already exists.
        # We avoid overwriting user-set fields.
        supabase.table("profiles").upsert(
            {"user_id": user_id},
            on_conflict="user_id"
        ).execute()

        return True, ""

    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def supa_list_my_drafts(sess, search_text: str = ""):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return [], err

    # Ensure the user has a profile row (safe, idempotent)
    ensure_profile_row(sess)

    try:
        supabase.postgrest.auth(access_token)

        res = (
            supabase.table("drafts")
            .select("id,title,subject,updated_at,created_at")
            .eq("user_id", user_id)
            .order("updated_at", desc=True)
            .execute()
        )

        rows = res.data or []

        s = (search_text or "").strip().lower()
        if s:
            rows = [
                r for r in rows
                if s in (r.get("title") or "").lower()
                or s in (r.get("subject") or "").lower()
            ]

        def _pretty_dt(iso_text: str) -> str:
            raw = (iso_text or "").strip()
            if not raw:
                return "Unknown date"
            try:
                raw = raw.replace("Z", "+00:00")
                dt_obj = datetime.fromisoformat(raw)
                return dt_obj.strftime("%d %b %Y")
            except Exception:
                short = raw[:10]
                return short if short else "Unknown date"

        def fmt_row(r):
            title = (r.get("title") or "Untitled draft").strip()
            subj = (r.get("subject") or "").strip()
            dt = _pretty_dt(r.get("updated_at") or r.get("created_at") or "")

            left = title
            if subj and subj != "—":
                left += f" — {subj}"

            # Keep the ID after | so parsing still works, but hide technical clutter from the visible part
            return f"{left} ({dt}) | {r['id']}"

        choices = [fmt_row(r) for r in rows]
        return choices, f"✅ Showing {len(choices)} drafts."

    except Exception as e:
        return [], f"❌ Failed to list drafts: {type(e).__name__}: {e}"


def supa_list_versions_for_draft(sess, draft_choice):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return [], err

    draft_id = _parse_draft_id(draft_choice)
    if not draft_id:
        return [], "No draft selected."

    try:
        res = (
            supabase.table("draft_versions")
            .select("version,created_at")
            .eq("draft_id", draft_id)
            .order("version", desc=True)
            .execute()
        )
        rows = getattr(res, "data", None) or []
        
        def _pretty_dt(iso_text: str) -> str:
            raw = (iso_text or "").strip()
            if not raw:
                return "Unknown date"
            try:
                raw = raw.replace("Z", "+00:00")
                dt_obj = datetime.fromisoformat(raw)
                return dt_obj.strftime("%d %b %Y")
            except Exception:
                short = raw[:10]
                return short if short else "Unknown date"

        choices = [
            f"Version {r['version']} — {_pretty_dt(r.get('created_at', ''))}"
            for r in rows
        ]
        return choices, f"✅ Found {len(choices)} versions."
    except Exception as e:
        return [], f"❌ Failed to list versions: {type(e).__name__}: {e}"


def supa_load_selected_version(sess, draft_choice, version_choice):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return "", "", err, "", 0

    draft_id = _parse_draft_id(draft_choice)
    if not draft_id:
        return "", "", "No draft selected.", "", 0

    v_num = _parse_version(version_choice)
    if not v_num:
        return "", "", "No version selected.", draft_id, 0

    try:
        res = (
            supabase.table("draft_versions")
            .select("doc_md,ppt_outline,version,created_at")
            .eq("draft_id", draft_id)
            .eq("version", v_num)
            .limit(1)
            .execute()
        )
        rows = getattr(res, "data", None) or []
        if not rows:
            return "", "", "❌ Version not found.", draft_id, 0

        row = rows[0]
        doc_md = row.get("doc_md") or ""
        ppt_outline = row.get("ppt_outline") or ""
        
        # Strip old visual-box HTML if it was saved in earlier versions
        doc_md = _strip_visual_box_html(doc_md)
        
        created_at = row.get("created_at") or ""
        return doc_md, ppt_outline, f"✅ Loaded {draft_id} v{v_num} ({created_at})", draft_id, v_num

    except Exception as e:
        return "", "", f"❌ Load failed: {type(e).__name__}: {e}", draft_id, 0


def _next_version_number(draft_id: str) -> int:
    res = (
        supabase.table("draft_versions")
        .select("version")
        .eq("draft_id", draft_id)
        .order("version", desc=True)
        .limit(1)
        .execute()
    )
    rows = getattr(res, "data", None) or []
    return (int(rows[0]["version"]) + 1) if rows else 1


def _supa_save_from_editor_core(
    sess,
    draft_name: str,
    subject: str,
    curriculum_stream: str,
    education_level: str,
    country: str,
    state_province: str,
    year_level: str,
    course: str,
    output_type:str,   # ✅ NEW
    edited_md: str,
    edited_ppt: str,
    current_draft_id: str,
):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return err, "", 0

    if not (edited_md or "").strip():
        return "Nothing to save: Markdown is empty.", current_draft_id or "", 0

    now = datetime.now(timezone.utc).isoformat()
    title = (draft_name or "").strip() or "Untitled draft"

    try:
        # NEW draft
        if not (current_draft_id or "").strip():
            draft_id = str(uuid.uuid4())
            draft_row = {
                "id": draft_id,
                "user_id": user_id,
                "title": title,
                "subject": (subject or "").strip() or None,
                "output_type": (output_type or "").strip() or None,
                "curriculum_stream": (curriculum_stream or "").strip() or None,
                "education_level": (education_level or "").strip() or None,
                "country": (country or "").strip() or None,
                "state_province": (state_province or "").strip() or None,
                "year_level": (year_level or "").strip() or None,
                "course": (course or "").strip() or None,
                "created_at": now,
                "updated_at": now,
            }
            supabase.table("drafts").insert(draft_row).execute()

            version_row = {
                "draft_id": draft_id,
                "version": 1,
                "doc_md": edited_md or "",
                "ppt_outline": edited_ppt or "",
                "created_at": now,
            }
            supabase.table("draft_versions").insert(version_row).execute()

            return f"✅ Saved NEW draft: {draft_id} (v1)", draft_id, 1

        # NEW version
        draft_id = current_draft_id.strip()
        v = _next_version_number(draft_id)

        supabase.table("drafts").update({
            "title": title,
            "subject": (subject or "").strip() or None,
            "output_type": (output_type or "").strip() or None,
            "curriculum_stream": (curriculum_stream or "").strip() or None,
            "education_level": (education_level or "").strip() or None,
            "country": (country or "").strip() or None,
            "state_province": (state_province or "").strip() or None,
            "year_level": (year_level or "").strip() or None,
            "course": (course or "").strip() or None,
            "updated_at": now,
        }).eq("id", draft_id).execute()

        version_row = {
            "draft_id": draft_id,
            "version": v,
            "doc_md": edited_md or "",
            "ppt_outline": edited_ppt or "",
            "created_at": now,
        }
        supabase.table("draft_versions").insert(version_row).execute()

        return f"✅ Saved new version: {draft_id} (v{v})", draft_id, v

    except Exception as e:
        return f"❌ Save failed: {type(e).__name__}: {e}", current_draft_id or "", 0


def supa_save_from_editor(
    sess,
    draft_name: str,
    subject: str,
    curriculum_stream: str,
    education_level: str,
    country: str,
    state_province: str,
    year_level: str,
    course: str,
    output_type: str,   # ✅ NEW
    edited_md: str,
    edited_ppt: str,
    current_draft_id: str,
):
    # Wrapper so UI can call either name safely
    return _supa_save_from_editor_core(
        sess,
        draft_name,
        subject,
        curriculum_stream,
        education_level,
        country,
        state_province,
        year_level,
        course,
        output_type,      # ✅ NEW
        edited_md,
        edited_ppt,
        current_draft_id,
    )


def supa_delete_version(sess, draft_choice, version_choice):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return err

    draft_id = _parse_draft_id(draft_choice)
    if not draft_id:
        return "No draft selected."

    v_num = _parse_version(version_choice)
    if not v_num:
        return "No version selected."

    try:
        supabase.table("draft_versions").delete().eq("draft_id", draft_id).eq("version", v_num).execute()

        res = supabase.table("draft_versions").select("version").eq("draft_id", draft_id).limit(1).execute()
        rows = getattr(res, "data", None) or []
        if not rows:
            supabase.table("drafts").delete().eq("id", draft_id).execute()
            return f"🗑️ Deleted v{v_num}. Draft had no versions left, so draft was deleted too."

        return f"🗑️ Deleted version v{v_num}."
    except Exception as e:
        return f"❌ Delete version failed: {type(e).__name__}: {e}"


def supa_delete_draft(sess, draft_choice):
    access_token, refresh_token, user_id, err = _require_session(sess)
    if err:
        return err

    draft_id = _parse_draft_id(draft_choice)
    if not draft_id:
        return "No draft selected."

    try:
        supabase.table("draft_versions").delete().eq("draft_id", draft_id).execute()
        supabase.table("drafts").delete().eq("id", draft_id).execute()
        return f"🔥 Deleted draft {draft_id} (and all versions)."
    except Exception as e:
        return f"❌ Delete draft failed: {type(e).__name__}: {e}"


def ui_supa_versions(sess, draft_choice):
    choices, msg = supa_list_versions_for_draft(sess, draft_choice)
    value = choices[0] if choices else None
    return gr.update(choices=choices, value=value), msg


# =============================
# GRADIO ACTIONS
# =============================
def action_generate_draft(
    supabase_session,
    input_mode, audio_path, typed_prompt, live_transcript,
    education_level,
    country, state_province,
    uni_country, university_name, faculty, module_code,
    year_level, course, other_subject, course_stream, output_type,
    include_memo, model_name,
    draft_mode, current_draft_id, edited_md_current
):
    try:
        can_proceed, limit_msg = check_rate_limit(supabase_session, action="generate")
        if not can_proceed:
            return "", "", "", None, limit_msg, "", input_mode, "", current_draft_id or "", 0

        if input_mode == "Speak (microphone)":
            lt = (live_transcript or "").strip()

            # ✅ Guard: if transcription is empty OR is an error message, do NOT generate
            if (not lt) or lt.lstrip().startswith("❌"):
                msg = lt if lt.lstrip().startswith("❌") else "Please record your instruction first."
                return "", "", "", None, f"{msg} {limit_msg}", "", input_mode, "", current_draft_id or "", 0

            instruction_text = lt
            transcript_out = instruction_text
            mode_used = "Speak (microphone)"

        else:
            tp = (typed_prompt or "").strip()
            if not tp:
                return "", "", "", None, f"Please type your instruction first. {limit_msg}", "", input_mode, "", current_draft_id or "", 0

            instruction_text = tp
            transcript_out = tp  # ✅ show the actual typed instruction (not a placeholder line)
            mode_used = "Type (keyboard)"

        is_edit = (draft_mode == "Edit current draft")

        if education_level == "University / Tertiary":
            eff_country = (uni_country or "").strip() or "Not specified"
            eff_state = ""
        else:
            eff_country = (country or "").strip() or "Not specified"
            eff_state = (state_province or "").strip()

        if education_level == "School (Primary / Secondary)":
            chosen = (course or "").strip()
        
            if chosen in {"— Choose subject —", "Choose Subject", ""}:
                effective_course = "Not specified"
            elif chosen == "Other (type it)":
                effective_course = (other_subject or "").strip() or "Not specified"
            else:
                effective_course = chosen
        else:
            effective_course = (module_code or "").strip() or "Not specified"
        
        curriculum_stream = (course_stream or "").strip()

        prompt = build_user_request(
            instruction_text=instruction_text,
            education_level=education_level,
            country=eff_country,
            state_province=eff_state,
            year_level=year_level,
            course=effective_course,
            output_type=output_type,
            include_memo=include_memo,
            curriculum_stream=curriculum_stream,
            university_name=university_name,
            faculty=faculty,
            module_code=module_code)

        if is_edit:
            if not (current_draft_id or "").strip():
                return transcript_out, "", "", None, "Edit mode selected, but no draft is loaded. Load a draft first in Draft Library.", instruction_text, mode_used, transcript_out, "", 0
            existing_md = (edited_md_current or "").strip()
            if not existing_md:
                return transcript_out, "", "", None, "Edit mode selected, but the editor is empty. Load a draft/version first.", instruction_text, mode_used, transcript_out, current_draft_id, 0
            prompt = (
                prompt
                + "\n\nEXISTING DOCUMENT (the teacher is editing this exact draft):\n<<<\n"
                + existing_md
                + "\n>>>\n"
            )

        llm_text = call_llm(prompt, model_name, edit_mode=is_edit)
        doc_md, ppt_outline, answer_key = split_sections(llm_text)
        combined_md = combine_doc_and_memo(doc_md, answer_key, include_memo)

        tmpdir = tempfile.mkdtemp()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        html_path = os.path.join(tmpdir, f"Preview_{stamp}.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(build_mathjax_html(combined_md))

        status = f"✅ Draft generated. {limit_msg}\nEdit it below, then Save (creates a Supabase version) or Export."
        return (
            transcript_out,
            combined_md,
            ppt_outline,
            html_path,
            status,
            instruction_text,
            mode_used,
            transcript_out,
            current_draft_id or "",
            0
        )

    except Exception as e:
        return "", "", "", None, safe_err("Draft generation failed.", e), "", input_mode, "", current_draft_id or "", 0


def action_export_files(sess, export_confirm: bool, edited_markdown: str, edited_ppt_outline: str):
    try:
        if not (edited_markdown or "").strip():
            return None, None, "Nothing to export: Markdown is empty."

        # --- V3.1: Only require confirmation for PAID AI images (kind='image') AND only when image-gen is enabled ---
        try:
            from config import ENABLE_IMAGE_GEN
            from credits import credits_needed_for_markdown

            if ENABLE_IMAGE_GEN:
                blob = (edited_markdown or "") + "\n\n" + (edited_ppt_outline or "")
                img_count, required = credits_needed_for_markdown(blob)

                # Confirmation is required ONLY if paid AI images exist
                if img_count > 0 and required > 0 and not export_confirm:
                    return None, None, (
                        "❌ Please tick the confirmation checkbox first "
                        "(export may spend credits on AI image generation)."
                    )
        except Exception:
            # Never block export due to warning logic failures
            pass

        # Enable per-image spending during THIS export
        def _spend(cost, reason, meta):
            return _spend_credits_for_session(sess, cost, reason, meta)

        # Inject the spender for this export run only
        diagram_library.set_credit_spend_fn(_spend)

        tmpdir = tempfile.mkdtemp()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        docx_path = os.path.join(tmpdir, f"EduDraft_Studio_{stamp}.docx")
        pptx_path = os.path.join(tmpdir, f"EduDraft_Studio_{stamp}.pptx")

        # -----------------------------
        # Render visuals ONCE, reuse for both DOCX + PPTX
        # -----------------------------
        md_clean = _strip_bundle_sections_for_docx(edited_markdown)
        md_clean = _strip_outer_md_fence(md_clean)
        md_norm = normalize_math_delimiters_for_pandoc(md_clean)

        try:
            md_rendered, generated_imgs = render_visuals_for_export(md_norm)
        except RuntimeError as e:
            if "Not enough credits to generate images/diagrams." in str(e):
                return None, None, "Not enough credits to generate images/diagrams."
            raise

        # DOCX: skip re-rendering visuals (prevents double spend)
        md_to_docx_with_editable_equations(md_rendered, docx_path, pre_rendered=True)

        outline = (edited_ppt_outline or "").strip()

        # If outline is empty OR contains no visuals, append an auto Visuals slide
        has_visuals_in_outline = ("[[VISUAL" in outline) or ("![" in outline)
        has_visuals_in_doc = bool(generated_imgs)

        if not outline:
            outline = "Slide 1: Lesson\n- (No outline provided)\nSpeaker notes: "

        if has_visuals_in_doc and not has_visuals_in_outline:
            nums = [int(n) for n in re.findall(r"Slide\s+(\d+):", outline)]
            next_n = (max(nums) + 1) if nums else 2

            visuals_lines = "\n".join([f"![]({p})" for p in generated_imgs if p])
            outline = outline.rstrip() + f"\n\nSlide {next_n}: Visuals\n{visuals_lines}\nSpeaker notes: "

        outline_to_pptx_with_math(outline, pptx_path)

        # ✅ IMPORTANT: return EXACTLY 3 outputs (docx, pptx, status)
        return (
            docx_path,
            pptx_path,
            "✅ Export complete. Scroll down to the bottom to download your DOCX and PPTX drafts."
        )

    except Exception as e:
        return None, None, safe_err("Export failed.", e)

    finally:
        # Critical: prevent cross-user credit spending in multi-user Gradio
        diagram_library.set_credit_spend_fn(None)
        

def toggle_input_visibility(mode):
    if mode == "Speak (microphone)":
        return gr.update(visible=True), gr.update(visible=False)
    else:
        return gr.update(visible=False), gr.update(visible=True, autofocus=True)


def toggle_education_visibility(level):
    is_school = (level == "School (Primary / Secondary)")
    return (
        gr.update(visible=is_school),
        gr.update(visible=is_school),
        gr.update(visible=not is_school),
        gr.update(visible=not is_school)
    )
    
def is_subject_ready(subject_choice: str, other_subject_text: str) -> bool:
    subject_choice = (subject_choice or "").strip()

    if not subject_choice:
        return False

    if subject_choice == "— Choose subject —":
        return False

    if subject_choice == "Other (type it)":
        return bool((other_subject_text or "").strip())

    return True


def toggle_other_subject(subject_choice: str):
    is_other = (subject_choice == "Other (type it)")
    # Show the textbox only when "Other" is selected
    return gr.update(visible=is_other)


def can_generate(mode, audio_path, typed_text):
    if mode == "Speak (microphone)":
        return audio_path is not None
    return bool((typed_text or "").strip())


def can_export(md_text):
    return bool((md_text or "").strip())
    

def can_save(md_text, sess):
    return bool((md_text or "").strip()) and (sess is not None)


def refresh_templates_on_open(session_state, category_filter):
    choices, status_msg = list_my_templates(session_state, category_filter)
    return gr.update(choices=choices, value=(choices[0] if choices else None)), status_msg


def _compute_generate_ui_state(input_mode, audio_path, typed_text, live_transcript, sess, subject_choice, other_subject_text, education_level):
    """
    Returns: gr.update(interactive=bool)
    Hard-lock overrides everything (paused / deletion requested).
    For School: requires subject + instruction readiness.
    For University: requires ONLY instruction readiness (no school subject).
    """
    locked, _reason = _account_lock_status(sess)
    if locked:
        return gr.update(interactive=False)

    is_school = (education_level == "School (Primary / Secondary)")

    # Subject must be valid ONLY for school mode
    if is_school and (not is_subject_ready(subject_choice, other_subject_text)):
        return gr.update(interactive=False)

    # Instruction readiness
    if input_mode == "Speak (microphone)":
        lt = (live_transcript or "").strip()
        ok = (audio_path is not None) and bool(lt) and (not lt.lstrip().startswith("❌"))
    else:
        ok = bool((typed_text or "").strip())

    return gr.update(interactive=ok)


# =============================
# EXPORT UI STATE MANAGEMENT (V3.1 FIXED) - DEBOUNCED VERSION
# =============================
def _compute_export_ui_state(md_text, ppt_text, snoozed, confirmed, sess):
    """
    SINGLE SOURCE OF TRUTH for export UI state.
    Returns: (note_html, note_visible, checkbox_visible, export_enabled, save_enabled)
    """
    from config import ENABLE_IMAGE_GEN
    from credits import credits_needed_for_markdown
    
    md_text = md_text or ""
    ppt_text = ppt_text or ""
    blob = md_text + "\n\n" + ppt_text
    
    has_doc = bool(md_text.strip())

    # ---- HARD LOCK: paused OR delete requested -> disable generate/save/export ----
    locked, lock_reason = _account_lock_status(sess)
    if locked:
        note = (
            f"## {lock_reason}\n\n"
            "Your account is locked. Generate, Save, and Export are disabled until this is resolved."
        )
        note_visible = True
        checkbox_visible = False
        export_enabled = False
        save_enabled = False
        return note, note_visible, checkbox_visible, export_enabled, save_enabled
    
    # No document -> disable everything
    if not has_doc:
        note = ""
        note_visible = False
        checkbox_visible = False
        export_enabled = False
        save_enabled = False
        return note, note_visible, checkbox_visible, export_enabled, save_enabled


    # ---- PLAN GATE: Free users can preview in markdown, but cannot Save/Export ----
    if not _is_pro_user(sess):
        note = (
            "🔒 **Pro required**\n\n"
            "You can still generate and preview your work in markdown on the Free plan.\n\n"
            "**To save versions or export DOCX/PPTX (including diagrams and AI images), please upgrade to Pro.**"
        )
        note_visible = True
        checkbox_visible = False
        export_enabled = False
        save_enabled = False
        return note, note_visible, checkbox_visible, export_enabled, save_enabled


    # If image-gen is OFF, no guard needed
    if not ENABLE_IMAGE_GEN:
        note = ""
        note_visible = False
        checkbox_visible = False
        export_enabled = True
        save_enabled = True
        return note, note_visible, checkbox_visible, export_enabled, save_enabled
    
    # Detect paid AI images
    try:
        img_count, required = credits_needed_for_markdown(blob)
        needs_paid_confirm = (img_count > 0 and required > 0)
    except Exception:
        needs_paid_confirm = False
    
    # No paid images -> no guard
    if not needs_paid_confirm:
        note = ""
        note_visible = False
        checkbox_visible = False
        export_enabled = True
        save_enabled = True
        return note, note_visible, checkbox_visible, export_enabled, save_enabled
    
    # Paid images exist but guard snoozed
    if snoozed:
        note = ""
        note_visible = False
        checkbox_visible = False
        export_enabled = True
        save_enabled = True
        return note, note_visible, checkbox_visible, export_enabled, save_enabled
    
    # Paid images + guard active + checkbox ticked
    if confirmed:
        note = ""
        note_visible = False
        checkbox_visible = True   # ✅ keep it visible while working
        export_enabled = True
        save_enabled = True
        return note, note_visible, checkbox_visible, export_enabled, save_enabled

    
    # Paid images + guard active + checkbox NOT ticked
    # Get balance to show in warning
    bal_text = ""
    if sess is not None:
        bal, err = get_balance(sess)
        if not err:
            bal_text = f" Your balance: {bal:.0f} credits."
    
    note = f"""⚠️ **Export Warning**

This document contains {img_count} AI-generated image(s) that will cost {required:.0f} credits to regenerate.{bal_text}

Please tick the checkbox below to enable export."""
    
    note_visible = True
    checkbox_visible = True
    export_enabled = False
    save_enabled = True   # Saving is always allowed; only export is gated
    
    return note, note_visible, checkbox_visible, export_enabled, save_enabled


def _update_export_ui(md_text, ppt_text, snoozed, confirmed, sess):
    """
    Wrapper to convert state tuple to Gradio updates.
    Returns: (export_cost_note, export_confirm, export_btn, save_version_btn)
    """
    note, note_visible, checkbox_visible, export_enabled, save_enabled = _compute_export_ui_state(
        md_text, ppt_text, snoozed, confirmed, sess
    )
    
    return (
        gr.update(value=note, visible=note_visible),
        gr.update(visible=checkbox_visible, value=confirmed),
        gr.update(interactive=export_enabled),
        gr.update(interactive=save_enabled)
    )


def _on_checkbox_change(md_text, ppt_text, snoozed, confirmed, sess):
    """
    Special handler for checkbox changes ONLY.
    This prevents the checkbox from disappearing when ticked.
    Returns: (export_btn, save_version_btn)
    """
    note, note_visible, checkbox_visible, export_enabled, save_enabled = _compute_export_ui_state(
        md_text, ppt_text, snoozed, confirmed, sess
    )
    
    # IMPORTANT: When checkbox is ticked, we DON'T hide it immediately
    # We only update the button states
    return (
        gr.update(interactive=export_enabled),
        gr.update(interactive=save_enabled)
    )


def _reset_export_guard():
    """Reset export guard state after new content is loaded/generated."""
    return False, False


def _after_export_snooze(status_text):
    t = (status_text or "").lower().strip()

    fail_indicators = [
        "export failed",
        "error:",
        "traceback",
        "exception",
        "❌",
        "please tick",
        "confirmation checkbox",
    ]
    if any(x in t for x in fail_indicators):
        # don't snooze; keep checkbox unticked
        return False, False

    success_indicators = [
        "✅", "✔", "✓",
        "export complete",
        "exported successfully",
        "export successful",
        "scroll down to download",
        "download your docx",
        "download your pptx",
    ]
    ok = any(x in t for x in success_indicators)

    if ok:
        # snooze warning; also reset confirm state
        return True, False

    return False, False


def action_live_transcribe(supabase_session, audio_path):
    try:
        if not audio_path:
            return ""

        # Optional (recommended): enforce your existing daily transcribe limit
        can_proceed, limit_msg = check_rate_limit(supabase_session, action="transcribe")
        if not can_proceed:
            return f"❌ {limit_msg}"

        text = (transcribe_audio(audio_path) or "").strip()
        return text or "❌ No speech detected / transcription came back empty. Try again."
    except Exception as e:
        return safe_err("Live transcription failed.", e)


# =============================
# UI
# =============================
COUNTRIES = sorted([c.name for c in pycountry.countries])

with gr.Blocks(title="EduDRAFT STUDIO", css=CUSTOM_CSS) as demo:

    supabase_session = gr.State(None)
    current_draft_id = gr.State("")
    current_version = gr.State(0)
    draft_subject_state = gr.State("")
    export_snooze = gr.State(False)  # hide export warning+checkbox after successful export until new draft/load
    export_confirm_state = gr.State(False)
    shared_year_level = gr.State("Year 11")
    shared_subject = gr.State("")



    last_instruction = gr.State("")
    last_mode = gr.State("Speak (microphone)")
    last_transcript = gr.State("")

    current_tab = gr.State("Login")
    is_logged_in = gr.State(False)
    browser_session_bridge = gr.Textbox(value="", visible=False, label="browser_session_bridge")

    gr.Markdown(
        "## 🎓 EduDRAFT STUDIO\n"
        "Welcome → Generate draft → Edit → Save versions → Load older versions → Export DOCX/PPTX.\n"
    )

    # =========================
    # GLOBAL HEADER (Facebook-style)
    # =========================
    global_banner_visible = gr.State(False)

    with gr.Group(visible=False, elem_id="global_banner") as global_banner:
        with gr.Row():
            global_avatar = gr.Image(label="", height=56, width=56, elem_id="global_avatar_img")
            with gr.Column():
                global_name = gr.Markdown("**(not loaded)**")
                global_signed = gr.Markdown("_Not signed in._")

    with gr.Tabs() as tabs:

        # =========================
        # LOGIN TAB
        # =========================
        with gr.TabItem("👋 Welcome", id="Login", visible=True) as login_tab:
            gr.Markdown("## 👋 Welcome")
            gr.Markdown("### Welcome to your digital office.")
            gr.Markdown("Choose what you want to do:")

            with gr.Row():
                go_signup = gr.Button("✍️ Create an account", scale=1)
                go_login  = gr.Button("🔐 I already have an account", scale=1)

            go_signup.click(fn=lambda: gr.update(selected="SignUp"), inputs=[], outputs=[tabs])
            go_login.click(fn=lambda: gr.update(selected="LogIn"), inputs=[], outputs=[tabs])

            # Keep old combined login UI below for now, but disable it
            if False:

                gr.Markdown("## 🔐 Account (Supabase)")

                with gr.Row():
                    auth_email = gr.Textbox(label="Email", placeholder="teacher@example.com")
                    auth_password = gr.Textbox(label="Password", type="password", placeholder="••••••••")

                gr.Markdown("### 📜 Legal (required to continue)")

                legal_privacy_cb = gr.Checkbox(label=f"I agree to the Privacy Policy ({LEGAL_VERSIONS['privacy']})", value=False)
                legal_terms_cb   = gr.Checkbox(label=f"I agree to the Terms of Use ({LEGAL_VERSIONS['terms']})", value=False)
                legal_tcs_cb     = gr.Checkbox(label=f"I agree to the Terms & Conditions ({LEGAL_VERSIONS['tcs']})", value=False)

                with gr.Row():
                    signup_btn = gr.Button("Sign up")
                    login_btn = gr.Button("Log in")
                    logout_btn = gr.Button("Log out")

                auth_status = gr.Textbox(label="Auth status", lines=2)

                whoami_btn_login = gr.Button("Who am I?")
                whoami_out_login = gr.Textbox(label="Current user", lines=4)

                secrets_box = gr.Textbox(
                    label="Supabase secrets (masked)",
                    lines=2,
                    value=supabase_secrets_masked(),
                    interactive=False
                )

                signup_btn.click(
                    fn=auth_signup_with_legal,
                    inputs=[auth_email, auth_password, legal_privacy_cb, legal_terms_cb, legal_tcs_cb],
                    outputs=[auth_status, supabase_session]
                )

                logout_btn.click(
                    fn=auth_logout,
                    inputs=[supabase_session],
                    outputs=[auth_status, supabase_session]
                ).then(
                    fn=lambda: ("Login", False),
                    inputs=[],
                    outputs=[current_tab, is_logged_in]
                ).then(
                    fn=lambda tab: gr.update(selected=tab),
                    inputs=[current_tab],
                    outputs=[tabs]
                ).then(
                    fn=lambda: (gr.update(visible=False), None, "**(not loaded)**", "_Not signed in._"),
                    inputs=[],
                    outputs=[global_banner, global_avatar, global_name, global_signed]
                )

                whoami_btn_login.click(
                    fn=auth_whoami,
                    inputs=[supabase_session],
                    outputs=[whoami_out_login]
                )

        # =========================
        # HOME TAB (premium lobby)
        # =========================
        with gr.TabItem("Home", id="Home", visible=False) as home_tab:

            # --- Header ---
            gr.Markdown("## 👋 Welcome back")
            gr.Markdown("Let’s create something meaningful today.")

            # --- Primary Actions ---
            with gr.Row():

                with gr.Column():
                    with gr.Group():
                        gr.Markdown("### ✍️ Create New Resource")
                        gr.Markdown("Start from scratch in the Workspace")
                        home_create_btn = gr.Button("Open Workspace")

                with gr.Column():
                    with gr.Group():
                        gr.Markdown("### 📂 Continue Your Work")
                        gr.Markdown("Open and edit an existing draft")
                        home_continue_btn = gr.Button("Open Drafts")

                with gr.Column():
                    with gr.Group():
                        gr.Markdown("### 🧩 Use a Template")
                        gr.Markdown("Start from a ready-made structure")
                        home_template_btn = gr.Button("Browse Templates")

            with gr.Group():
                gr.Markdown("### 👤 Keep your profile up to date")
                gr.Markdown(
                    "Your profile helps personalise future output, including curriculum, teaching preferences, and account settings."
                )
                home_profile_btn = gr.Button("Open Profile")

            # --- System Status ---
            gr.Markdown("---")
            home_status = gr.Markdown(
                "🟢 Status: Logged in  \n📦 Plan: Free"
            )

            # --- Soft Guidance ---
            gr.Markdown(
                "_Use the menu above to explore all features._"
            )

        # --- Home Navigation ---
        home_create_btn.click(
            fn=lambda: "Workspace",
            outputs=[current_tab]
        ).then(
            fn=lambda tab: gr.update(selected=tab),
            inputs=[current_tab],
            outputs=[tabs]
        )

        home_continue_btn.click(
            fn=lambda: "Drafts",
            outputs=[current_tab]
        ).then(
            fn=lambda tab: gr.update(selected=tab),
            inputs=[current_tab],
            outputs=[tabs]
        )

        home_template_btn.click(
            fn=lambda: "Templates",
            outputs=[current_tab]
        ).then(
            fn=lambda tab: gr.update(selected=tab),
            inputs=[current_tab],
            outputs=[tabs]
        )

        home_profile_btn.click(
            fn=lambda: "Profile",
            outputs=[current_tab]
        ).then(
            fn=lambda tab: gr.update(selected=tab),
            inputs=[current_tab],
            outputs=[tabs]
        )

        # =========================
        # NOTIFICATIONS TAB
        # =========================
        with gr.TabItem("Notifications", id="Notifications", visible=False) as notifications_tab:
            gr.Markdown("## 🔔 Notifications & Notices")
            gr.Markdown(
                "Before you can continue into EduDRAFT STUDIO, you need to complete the required items below."
            )

            gr.Markdown("### Required Legal Acknowledgements")

            with gr.Accordion("🛡️ Privacy Policy", open=True) as notifications_privacy_acc:
                gr.Markdown(
                    "Covers how EduDRAFT STUDIO collects, stores, and protects your information."
                )
                notifications_privacy_doc_md = gr.Markdown(
                    value=LEGAL_MD_TEXT.get("privacy", ""),
                    elem_classes=["legal_doc_box"]
                )
                notifications_privacy_ack_status = gr.Markdown("❌ Not acknowledged for current version.")
                notifications_privacy_ack_cb = gr.Checkbox(
                    label="I have read and agree to the Privacy Policy",
                    value=False
                )

            with gr.Accordion("📘 Terms of Use", open=False) as notifications_terms_acc:
                gr.Markdown(
                    "Covers acceptable use of the app, teacher responsibilities, and content review expectations."
                )
                notifications_terms_doc_md = gr.Markdown(
                    value=LEGAL_MD_TEXT.get("terms", ""),
                    elem_classes=["legal_doc_box"]
                )
                notifications_terms_ack_status = gr.Markdown("❌ Not acknowledged for current version.")
                notifications_terms_ack_cb = gr.Checkbox(
                    label="I have read and agree to the Terms of Use",
                    value=False
                )

            with gr.Accordion("🧾 Terms & Conditions", open=False) as notifications_tcs_acc:
                gr.Markdown(
                    "Covers commercial terms for the service, including plan rules and credit-related terms."
                )
                notifications_tcs_doc_md = gr.Markdown(
                    value=LEGAL_MD_TEXT.get("tcs", ""),
                    elem_classes=["legal_doc_box"]
                )
                notifications_tcs_ack_status = gr.Markdown("❌ Not acknowledged for current version.")
                notifications_tcs_ack_cb = gr.Checkbox(
                    label="I have read and agree to the Terms & Conditions",
                    value=False
                )

        # =========================
        # SIGN UP TAB (placeholder)
        # =========================
        with gr.TabItem("Sign up", id="SignUp", visible=True) as signup_tab:
            gr.Markdown("## ✍️ Sign up")

            with gr.Row():
                signup_email = gr.Textbox(label="Email", placeholder="teacher@example.com")

            signup_showing_password = gr.State(False)

            with gr.Row():
                signup_password = gr.Textbox(
                    label="Password",
                    type="password",
                    placeholder="••••••••",
                    scale=8,
                    visible=True
                )
                signup_password_plain = gr.Textbox(
                    label="Password",
                    type="text",
                    placeholder="••••••••",
                    scale=8,
                    visible=False
                )
                signup_password_toggle = gr.Button("👁", min_width=55)

            signup_password_toggle.click(
                fn=_toggle_password_pair,
                inputs=[signup_password, signup_password_plain, signup_showing_password],
                outputs=[signup_password, signup_password_plain, signup_password_toggle, signup_showing_password]
            )

            gr.Markdown("### 📜 Legal (required to create an account)")
            gr.Markdown(
                "Please open and review each document before agreeing. "
                "Each document will open below, and the acknowledgement box appears at the bottom of that document only."
            )

            # --------------------------------------------------
            # SIGN-UP LEGAL UI STATE
            # --------------------------------------------------
            signup_legal_open_doc = gr.State("")

            # These remain the canonical booleans fed into auth_signup_with_legal(...)
            # They are now hidden from the user and driven by the document UI below.
            signup_legal_privacy = gr.Checkbox(
                label=f"I agree to the Privacy Policy ({LEGAL_VERSIONS['privacy']})",
                value=False,
                visible=False
            )
            signup_legal_terms = gr.Checkbox(
                label=f"I agree to the Terms of Use ({LEGAL_VERSIONS['terms']})",
                value=False,
                visible=False
            )
            signup_legal_tcs = gr.Checkbox(
                label=f"I agree to the Terms & Conditions ({LEGAL_VERSIONS['tcs']})",
                value=False,
                visible=False
            )

            def _signup_legal_open(doc_key, privacy_done, terms_done, tcs_done):
                doc_key = (doc_key or "").strip().lower()

                privacy_done = bool(privacy_done)
                terms_done = bool(terms_done)
                tcs_done = bool(tcs_done)

                def _summary(label, ver, done):
                    return f"✅ {label} acknowledged ({ver})" if done else f"📄 Read {label} ({ver})"

                privacy_summary = _summary("Privacy Policy", LEGAL_VERSIONS["privacy"], privacy_done)
                terms_summary   = _summary("Terms of Use", LEGAL_VERSIONS["terms"], terms_done)
                tcs_summary     = _summary("Terms & Conditions", LEGAL_VERSIONS["tcs"], tcs_done)

                privacy_show = (doc_key == "privacy" and not privacy_done)
                terms_show   = (doc_key == "terms" and not terms_done)
                tcs_show     = (doc_key == "tcs" and not tcs_done)

                return (
                    doc_key,
                    gr.update(value=privacy_summary),
                    gr.update(value=terms_summary),
                    gr.update(value=tcs_summary),
                    gr.update(visible=privacy_show),
                    gr.update(visible=terms_show),
                    gr.update(visible=tcs_show)
                )

            def _signup_legal_ack(doc_key, checked, privacy_done, terms_done, tcs_done):
                doc_key = (doc_key or "").strip().lower()
                checked = bool(checked)

                privacy_done = bool(privacy_done)
                terms_done = bool(terms_done)
                tcs_done = bool(tcs_done)

                if doc_key == "privacy" and checked:
                    privacy_done = True
                elif doc_key == "terms" and checked:
                    terms_done = True
                elif doc_key == "tcs" and checked:
                    tcs_done = True

                def _summary(label, ver, done):
                    return f"✅ {label} acknowledged ({ver})" if done else f"📄 Read {label} ({ver})"

                return (
                    privacy_done,
                    terms_done,
                    tcs_done,
                    "",
                    gr.update(value=_summary("Privacy Policy", LEGAL_VERSIONS["privacy"], privacy_done)),
                    gr.update(value=_summary("Terms of Use", LEGAL_VERSIONS["terms"], terms_done)),
                    gr.update(value=_summary("Terms & Conditions", LEGAL_VERSIONS["tcs"], tcs_done)),
                    gr.update(visible=False),
                    gr.update(visible=False),
                    gr.update(visible=False),
                    gr.update(value=False),
                    gr.update(value=False),
                    gr.update(value=False)
                )

            with gr.Group():
                signup_privacy_open_btn = gr.Button(
                    f"📄 Read Privacy Policy ({LEGAL_VERSIONS['privacy']})",
                    variant="secondary"
                )
                signup_privacy_panel = gr.Group(visible=False)
                with signup_privacy_panel:
                    gr.Markdown(
                        "Covers how EduDRAFT STUDIO collects, stores, and protects your information."
                    )
                    signup_privacy_doc_md = gr.Markdown(
                        value=LEGAL_MD_TEXT.get("privacy", ""),
                        elem_classes=["legal_doc_box"]
                    )
                    signup_privacy_ack_ui = gr.Checkbox(
                        label="I have read and agree to the Privacy Policy",
                        value=False
                    )

            with gr.Group():
                signup_terms_open_btn = gr.Button(
                    f"📄 Read Terms of Use ({LEGAL_VERSIONS['terms']})",
                    variant="secondary"
                )
                signup_terms_panel = gr.Group(visible=False)
                with signup_terms_panel:
                    gr.Markdown(
                        "Covers acceptable use of the app, teacher responsibilities, and content review expectations."
                    )
                    signup_terms_doc_md = gr.Markdown(
                        value=LEGAL_MD_TEXT.get("terms", ""),
                        elem_classes=["legal_doc_box"]
                    )
                    signup_terms_ack_ui = gr.Checkbox(
                        label="I have read and agree to the Terms of Use",
                        value=False
                    )

            with gr.Group():
                signup_tcs_open_btn = gr.Button(
                    f"📄 Read Terms & Conditions ({LEGAL_VERSIONS['tcs']})",
                    variant="secondary"
                )
                signup_tcs_panel = gr.Group(visible=False)
                with signup_tcs_panel:
                    gr.Markdown(
                        "Covers commercial terms for the service, including plan rules and credit-related terms."
                    )
                    signup_tcs_doc_md = gr.Markdown(
                        value=LEGAL_MD_TEXT.get("tcs", ""),
                        elem_classes=["legal_doc_box"]
                    )
                    signup_tcs_ack_ui = gr.Checkbox(
                        label="I have read and agree to the Terms & Conditions",
                        value=False
                    )

            signup_privacy_open_btn.click(
                fn=lambda pp, tou, tcs: _signup_legal_open("privacy", pp, tou, tcs),
                inputs=[signup_legal_privacy, signup_legal_terms, signup_legal_tcs],
                outputs=[
                    signup_legal_open_doc,
                    signup_privacy_open_btn,
                    signup_terms_open_btn,
                    signup_tcs_open_btn,
                    signup_privacy_panel,
                    signup_terms_panel,
                    signup_tcs_panel,
                ]
            )

            signup_terms_open_btn.click(
                fn=lambda pp, tou, tcs: _signup_legal_open("terms", pp, tou, tcs),
                inputs=[signup_legal_privacy, signup_legal_terms, signup_legal_tcs],
                outputs=[
                    signup_legal_open_doc,
                    signup_privacy_open_btn,
                    signup_terms_open_btn,
                    signup_tcs_open_btn,
                    signup_privacy_panel,
                    signup_terms_panel,
                    signup_tcs_panel,
                ]
            )

            signup_tcs_open_btn.click(
                fn=lambda pp, tou, tcs: _signup_legal_open("tcs", pp, tou, tcs),
                inputs=[signup_legal_privacy, signup_legal_terms, signup_legal_tcs],
                outputs=[
                    signup_legal_open_doc,
                    signup_privacy_open_btn,
                    signup_terms_open_btn,
                    signup_tcs_open_btn,
                    signup_privacy_panel,
                    signup_terms_panel,
                    signup_tcs_panel,
                ]
            )

            signup_privacy_ack_ui.input(
                fn=lambda checked, pp, tou, tcs: _signup_legal_ack("privacy", checked, pp, tou, tcs),
                inputs=[signup_privacy_ack_ui, signup_legal_privacy, signup_legal_terms, signup_legal_tcs],
                outputs=[
                    signup_legal_privacy,
                    signup_legal_terms,
                    signup_legal_tcs,
                    signup_legal_open_doc,
                    signup_privacy_open_btn,
                    signup_terms_open_btn,
                    signup_tcs_open_btn,
                    signup_privacy_panel,
                    signup_terms_panel,
                    signup_tcs_panel,
                    signup_privacy_ack_ui,
                    signup_terms_ack_ui,
                    signup_tcs_ack_ui,
                ]
            )

            signup_terms_ack_ui.input(
                fn=lambda checked, pp, tou, tcs: _signup_legal_ack("terms", checked, pp, tou, tcs),
                inputs=[signup_terms_ack_ui, signup_legal_privacy, signup_legal_terms, signup_legal_tcs],
                outputs=[
                    signup_legal_privacy,
                    signup_legal_terms,
                    signup_legal_tcs,
                    signup_legal_open_doc,
                    signup_privacy_open_btn,
                    signup_terms_open_btn,
                    signup_tcs_open_btn,
                    signup_privacy_panel,
                    signup_terms_panel,
                    signup_tcs_panel,
                    signup_privacy_ack_ui,
                    signup_terms_ack_ui,
                    signup_tcs_ack_ui,
                ]
            )

            signup_tcs_ack_ui.input(
                fn=lambda checked, pp, tou, tcs: _signup_legal_ack("tcs", checked, pp, tou, tcs),
                inputs=[signup_tcs_ack_ui, signup_legal_privacy, signup_legal_terms, signup_legal_tcs],
                outputs=[
                    signup_legal_privacy,
                    signup_legal_terms,
                    signup_legal_tcs,
                    signup_legal_open_doc,
                    signup_privacy_open_btn,
                    signup_terms_open_btn,
                    signup_tcs_open_btn,
                    signup_privacy_panel,
                    signup_terms_panel,
                    signup_tcs_panel,
                    signup_privacy_ack_ui,
                    signup_terms_ack_ui,
                    signup_tcs_ack_ui,
                ]
            )

            with gr.Row():
                signup_create_btn = gr.Button("Create account")
                signup_logout_btn = gr.Button("Log out")

            signup_status = gr.Textbox(label="Sign up status", lines=2)

            signup_whoami_btn = gr.Button("Who am I?")
            signup_whoami_out = gr.Textbox(label="Current user", lines=4)

            # Create account (legal required)
            signup_create_btn.click(
                fn=lambda email, pwd_masked, pwd_plain, showing, pp, terms, tcs: auth_signup_with_legal(
                    email,
                    (pwd_plain if showing else pwd_masked) or "",
                    pp,
                    terms,
                    tcs
                ),
                inputs=[signup_email, signup_password, signup_password_plain, signup_showing_password, signup_legal_privacy, signup_legal_terms, signup_legal_tcs],
                outputs=[signup_status, supabase_session]
            ).then(
                fn=lambda msg, sess: _tab_after_auth_or_stay("SignUp", msg, sess),
                inputs=[signup_status, supabase_session],
                outputs=[current_tab, is_logged_in]
            ).then(
                fn=lambda tab: gr.update(selected=tab),
                inputs=[current_tab],
                outputs=[tabs]
            ).then(
                fn=_session_to_browser_blob,
                inputs=[supabase_session],
                outputs=[browser_session_bridge]
            ).then(
                fn=None,
                inputs=[browser_session_bridge],
                outputs=[browser_session_bridge],
                queue=False,
                show_progress="hidden",
                js=f"""
                (blob) => {{
                    const key = "{BROWSER_SESSION_STORAGE_KEY}";
                    try {{
                        if (blob) {{
                            window.localStorage.setItem(key, blob);
                            return blob;
                        }}
                        return "";
                    }} catch (e) {{
                        return "";
                    }}
                }}
                """
            )

            signup_logout_btn.click(
                fn=auth_logout,
                inputs=[supabase_session],
                outputs=[signup_status, supabase_session]
            ).then(
                fn=lambda: ("Login", False),
                inputs=[],
                outputs=[current_tab, is_logged_in]
            ).then(
                fn=lambda tab: gr.update(selected=tab),
                inputs=[current_tab],
                outputs=[tabs]
            ).then(
                fn=None,
                inputs=[],
                outputs=[browser_session_bridge],
                queue=False,
                show_progress="hidden",
                js=f"""
                () => {{
                    const key = "{BROWSER_SESSION_STORAGE_KEY}";
                    try {{
                        window.localStorage.removeItem(key);
                    }} catch (e) {{}}
                    return "";
                }}
                """
            )

            signup_whoami_btn.click(
                fn=auth_whoami,
                inputs=[supabase_session],
                outputs=[signup_whoami_out]
            )

        def _clear_login_fields_after_auth(msg, sess):
            ok = bool(sess and isinstance(sess, dict) and sess.get("access_token"))
            if ok:
                return "", "", ""
            return gr.update(), gr.update(), gr.update()

        # =========================
        # LOG IN TAB (placeholder)
        # =========================
        with gr.TabItem("Log in", id="LogIn", visible=True) as login2_tab:
            gr.Markdown("## 🔐 Log in")

            with gr.Row():
                login_email = gr.Textbox(label="Email", placeholder="teacher@example.com")

            login_showing_password = gr.State(False)

            with gr.Row():
                login_password = gr.Textbox(
                    label="Password",
                    type="password",
                    placeholder="••••••••",
                    scale=8,
                    visible=True
                )
                login_password_plain = gr.Textbox(
                    label="Password",
                    type="text",
                    placeholder="••••••••",
                    scale=8,
                    visible=False
                )
                login_password_toggle = gr.Button("👁", min_width=55)

                login_password_toggle.click(
                    fn=_toggle_password_pair,
                    inputs=[login_password, login_password_plain, login_showing_password],
                    outputs=[login_password, login_password_plain, login_password_toggle, login_showing_password]
                )

            with gr.Row():
                login_go_btn = gr.Button("Log in")
                login_logout_btn = gr.Button("Log out")

            login_status = gr.Textbox(label="Auth status", lines=2)

            login_whoami_btn = gr.Button("Who am I?")
            login_whoami_out = gr.Textbox(label="Current user", lines=4)

            # Log in (fast path: no legal ticking here)
            login_go_btn.click(
                fn=lambda email, pwd_masked, pwd_plain, showing: auth_login_with_legal(
                    email,
                    (pwd_plain if showing else pwd_masked) or "",
                    False,
                    False,
                    False
                ),
                inputs=[login_email, login_password, login_password_plain, login_showing_password],
                outputs=[login_status, supabase_session]
            ).then(
                fn=lambda msg, sess: _tab_after_auth_or_stay("LogIn", msg, sess),
                inputs=[login_status, supabase_session],
                outputs=[current_tab, is_logged_in]
            ).then(
                fn=_clear_login_fields_after_auth,
                inputs=[login_status, supabase_session],
                outputs=[login_email, login_password, login_password_plain]
            ).then(
                fn=lambda tab: gr.update(selected=tab),
                inputs=[current_tab],
                outputs=[tabs]
            ).then(
                fn=_session_to_browser_blob,
                inputs=[supabase_session],
                outputs=[browser_session_bridge]
            ).then(
                fn=None,
                inputs=[browser_session_bridge],
                outputs=[browser_session_bridge],
                queue=False,
                show_progress="hidden",
                js=f"""
                (blob) => {{
                    const key = "{BROWSER_SESSION_STORAGE_KEY}";
                    try {{
                        if (blob) {{
                            window.localStorage.setItem(key, blob);
                            return blob;
                        }}
                        return "";
                    }} catch (e) {{
                        return "";
                    }}
                }}
                """
            )

            login_logout_btn.click(
                fn=auth_logout,
                inputs=[supabase_session],
                outputs=[login_status, supabase_session]
            ).then(
                fn=lambda: ("Login", False),
                inputs=[],
                outputs=[current_tab, is_logged_in]
            ).then(
                fn=lambda tab: gr.update(selected=tab),
                inputs=[current_tab],
                outputs=[tabs]
            ).then(
                fn=lambda: (gr.update(visible=False), None, "**(not loaded)**", "_Not signed in._"),
                inputs=[],
                outputs=[global_banner, global_avatar, global_name, global_signed]
            ).then(
                fn=None,
                inputs=[],
                outputs=[browser_session_bridge],
                queue=False,
                show_progress="hidden",
                js=f"""
                () => {{
                    const key = "{BROWSER_SESSION_STORAGE_KEY}";
                    try {{
                        window.localStorage.removeItem(key);
                    }} catch (e) {{}}
                    return "";
                }}
                """
            )

            login_whoami_btn.click(
                fn=auth_whoami,
                inputs=[supabase_session],
                outputs=[login_whoami_out]
            )


        # =========================
        # WORKSPACE TAB
        # =========================
        with gr.TabItem("Workspace", id="Workspace", visible=False) as workspace_tab:
            gr.Markdown("## ✍️ Workspace")
            credits_refresh_state_ws = gr.Textbox(visible=False)

            # --- Teacher-first PDF panel (hidden until needed) ---
            with gr.Row():
                start_from_choice = gr.Dropdown(
                    choices=[
                        ("Choose…", ""),
                        ("📄 Start from PDF", "PDF"),
                        ("📝 Start from DOCX", "DOCX"),
                        ("📊 Start from PPTX", "PPTX"),
                    ],
                    value="",
                    label="Start from…")
                hide_upload_panel_btn = gr.Button("✖ Hide upload panel")

            pdf_panel = gr.Group(visible=False)
            with pdf_panel:
                gr.Markdown(
                    "**Upload an existing document and continue working on it in the editor.**\n\n"
                    "This is for when you already have a test, worksheet, memo, or exam paper and want to edit, improve, or reuse it.\n\n"
                    "- The file will be converted into an editable draft\n"
                    "- This is NOT a template — it becomes working content in your Workspace"
                )
                pdf_upload = gr.File(label="Upload PDF", file_types=[".pdf"])
                pdf_draft_title = gr.Textbox(
                    label="Draft name (optional)",
                    placeholder="Leave blank to use the PDF filename"
                )
                pdf_generate_btn = gr.Button(
                    "📄 Generate Draft from PDF (Auto-save v1)",
                    variant="primary",
                    interactive=False
                )
                pdf_status_hint = gr.Markdown(
                    "Tip: If the PDF is scanned (image-only), text extraction may fail. (OCR can be added later.)"
                )

            # --- Create / Edit Controls ---
            input_mode = gr.Radio(
                ["Speak (microphone)", "Type (keyboard)"],
                value="Speak (microphone)",
                label="How would you like to give the instruction?"
            )
            
            docx_panel = gr.Group(visible=False)
            with docx_panel:
                gr.Markdown(
                    "**Upload a DOCX file and choose how EduDRAFT should use it.**\n\n"
                    "Use **Convert to editable draft** when you want to keep working on the same document.\n\n"
                    "Use **Generate similar new paper** when you want EduDRAFT to treat the uploaded exam/test as a model paper and create a new equivalent version with different content."
                )

                docx_mode = gr.Radio(
                    choices=[
                        "Convert to editable draft",
                        "Generate similar new paper"
                    ],
                    value="Convert to editable draft",
                    label="DOCX mode"
                )

                docx_upload = gr.File(label="Upload DOCX", file_types=[".docx"])

                docx_draft_title = gr.Textbox(
                    label="Draft name (optional)",
                    placeholder="Leave blank to use the DOCX filename"
                )

                docx_generate_btn = gr.Button(
                    "📝 Generate from DOCX (Auto-save v1)",
                    variant="primary",
                    interactive=False
                )

                docx_status_hint = gr.Markdown(
                    "Tip: Convert mode is text-first and may simplify images/tables. Similar-paper mode will preserve the model structure as a blueprint."
                )

            pptx_panel = gr.Group(visible=False)
            with pptx_panel:
                gr.Markdown(
                    "**Upload a PPTX and convert it into editable content.**\n\n"
                    "Useful if you want to reuse slides as structured teaching content.\n\n"
                    "- Converted into a draft in the Workspace\n"
                    "- Not saved as a template"
                )
                pptx_upload = gr.File(label="Upload PPTX", file_types=[".pptx"])
                pptx_draft_title = gr.Textbox(
                    label="Draft name (optional)",
                    placeholder="Leave blank to use the PPTX filename"
                )
                pptx_generate_btn = gr.Button(
                    "📊 Generate Draft from PPTX (Auto-save v1)",
                    variant="primary",
                    interactive=False
                )
                pptx_status_hint = gr.Markdown(
                    "Tip: If the PPTX is mostly images, text extraction may be low (OCR/images later).")
                
            # =========================
            # Upload panel visibility control
            # =========================
            
            def _show_upload_panel(choice: str):
                choice = (choice or "None").strip()
                return (
                    gr.update(visible=(choice == "PDF")),
                    gr.update(visible=(choice == "DOCX")),
                    gr.update(visible=(choice == "PPTX")))
            
            def _hide_all_upload_panels():
                return (
                    gr.update(visible=False),
                    gr.update(visible=False),
                    gr.update(visible=False),
                    gr.update(value=""))
            
            start_from_choice.change(
                fn=_show_upload_panel,
                inputs=[start_from_choice],
                outputs=[pdf_panel, docx_panel, pptx_panel]
            )
            
            hide_upload_panel_btn.click(
                fn=_hide_all_upload_panels,
                inputs=[],
                outputs=[pdf_panel, docx_panel, pptx_panel, start_from_choice]
            )

            def _on_pdf_upload(pdf_file, subj, other_subj, ed_level):
                # 1) Run the analyzer (keeps your “scanned pdf / extraction” messaging)
                hint, _btn_update = action_analyze_pdf(pdf_file)

                # 2) Apply the subject gate (school requires subject; uni does not)
                interactive = False
                if pdf_file is not None:
                    if ed_level == "University / Tertiary":
                        interactive = True
                    else:
                        interactive = is_subject_ready(subj, other_subj)

                return hint, gr.update(interactive=interactive)

            docx_upload.change(
                fn=action_analyze_docx,
                inputs=[docx_upload],
                outputs=[docx_status_hint, docx_generate_btn]
            )

            pptx_upload.change(
                fn=action_analyze_pptx,
                inputs=[pptx_upload],
                outputs=[pptx_status_hint, pptx_generate_btn]
            )

            audio_in = gr.Audio(
                sources=["microphone"],
                type="filepath",
                label="Microphone input",
                visible=True
            )

            live_transcript = gr.Textbox(
                label="Live transcript (editable) — generated after recording",
                lines=3,
                placeholder="Record audio… transcription will appear here. You can edit it before generating.",
                visible=True
            )

            audio_in.change(
                fn=action_live_transcribe,
                inputs=[supabase_session, audio_in],
                outputs=[live_transcript]
            )

            typed_prompt = gr.Textbox(
                label="Typed instruction",
                placeholder="Example: Create a Year 11 Methods test on differentiation. 45 minutes. 40 marks. Include chain rule + product rule. Include memo.",
                lines=5,
                visible=False
            )

            input_mode.change(fn=toggle_input_visibility, inputs=input_mode, outputs=[audio_in, typed_prompt])

            education_level = gr.Radio(
                ["School (Primary / Secondary)", "University / Tertiary"],
                value="School (Primary / Secondary)",
                label="Education level"
            )

            with gr.Row(visible=True) as school_row:
                country = gr.Dropdown(choices=COUNTRIES, value="Australia", label="Country (type to search)")
                state_province = gr.Textbox(
                    label="State / Province / Region (optional)",
                    placeholder="e.g. Western Australia, Texas, Ontario, Gauteng"
                )

            with gr.Row(visible=True) as school_row2:
                year_level = gr.Dropdown(
                    ["Year 6", "Year 7", "Year 8", "Year 9", "Year 10", "Year 11", "Year 12"],
                    value="Year 11",
                    label="Year level"
                )
                year_level.change(
                    fn=lambda v: v,
                    inputs=[year_level],
                    outputs=[shared_year_level]
                )
                course = gr.Dropdown(
                    choices=[
                        "— Choose subject —",
                        "Other (type it)",
                        "English",
                        "Mathematics",
                        "Mathematics (Methods)",
                        "Mathematics (Applications)",
                        "Mathematics (General)",
                        "Science",
                        "Biology",
                        "Chemistry",
                        "Physics",
                        "Human Biology",
                        "HASS / Humanities",
                        "Geography",
                        "History",
                        "Economics",
                        "Accounting & Finance",
                        "Business / Commerce",
                        "Digital Technologies / Computing",
                        "Design & Technology",
                        "Health & Physical Education",
                        "Drama",
                        "Music",
                        "Visual Arts",
                        "Languages",
                    ],
                    value="— Choose subject —",
                    label="Subject * (required)")
                other_subject = gr.Textbox(
                    label="Other subject (required)",
                    placeholder="Type the subject name here…",
                    visible=False
                )
                def _subject_to_store(subj, other_subj):
                    subj = (subj or "").strip()
                
                    if subj == "— Choose subject —":
                        return ""
                
                    if subj == "Other (type it)":
                        return (other_subj or "").strip()
                
                    return subj

                course.change(
                    fn=_subject_to_store,
                    inputs=[course, other_subject],
                    outputs=[draft_subject_state]
                )

                draft_subject_state.change(
                    fn=lambda v: v,
                    inputs=[draft_subject_state],
                    outputs=[shared_subject]
                )

                other_subject.change(
                    fn=_subject_to_store,
                    inputs=[course, other_subject],
                    outputs=[draft_subject_state]
                )
                
                cancel_other_subject_btn = gr.Button("✖ Cancel 'Other' and go back to list", visible=False)

                course_stream = gr.Textbox(
                    label="Curriculum / syllabus (optional)",
                    placeholder="e.g. WA ATAR, IGCSE, A-Level, IB AA/AI, CAPS, NSC, GCSE…",
                    visible=True
                )

            course.change(
                fn=lambda s: (toggle_other_subject(s), gr.update(visible=(s == "Other (type it)"))),
                inputs=course,
                outputs=[other_subject, cancel_other_subject_btn]
            )

            cancel_other_subject_btn.click(
                fn=lambda: (gr.update(value="Mathematics (Methods)"), gr.update(value="", visible=False), gr.update(visible=False)),
                inputs=[],
                outputs=[course, other_subject, cancel_other_subject_btn]
            )

            with gr.Row(visible=False) as uni_row:
                uni_country = gr.Dropdown(choices=COUNTRIES, value="Australia", label="Country (type to search)")
                university_name = gr.Textbox(label="University / Institution", placeholder="e.g. UWA, MIT, UCT")

            with gr.Row(visible=False) as uni_row2:
                faculty = gr.Textbox(label="Faculty / Discipline (optional)", placeholder="e.g. Mathematics, Engineering")
                module_code = gr.Textbox(label="Course / Module (optional)", placeholder="e.g. MATH1012 Calculus")

            education_level.change(
                fn=toggle_education_visibility,
                inputs=education_level,
                outputs=[school_row, school_row2, uni_row, uni_row2]
            )

            pdf_upload.change(
                fn=_on_pdf_upload,
                inputs=[pdf_upload, course, other_subject, education_level],
                outputs=[pdf_status_hint, pdf_generate_btn]
            )

            output_type = gr.Dropdown(
                ["Test / Quiz", "Worksheet", "Investigation", "Lesson", "PowerPoint lesson", "Marking key / Memo only", "Rubric", "Homework", "Exam", "Lesson Plan", "Revision Sheet", "Custom"],
                value="Test / Quiz",
                label="Output type"
            )

            with gr.Row():
                include_memo = gr.Checkbox(value=True, label="Include memo / answer key (when relevant)")
                model_name = gr.Dropdown(["gpt-5.2", "gpt-4o", "gpt-4o-mini"], value="gpt-5.2", label="Model")

            draft_mode = gr.Radio(
                ["Create new draft", "Edit current draft"],
                value="Create new draft",
                label="Draft mode"
            )

            # ==================================
            # Export area (button + credit guard) - FIXED
            # ==================================
            with gr.Group(elem_id="export_area"):

                with gr.Row(elem_id="action_buttons_row"):
                    gen_btn = gr.Button("Generate Draft", interactive=False)
                    save_version_btn = gr.Button("💾 Save NEW version", interactive=False)
                    export_btn = gr.Button("📦 Export DOCX & PPTX", interactive=False, elem_id="export_btn", elem_classes=["export_btn"])

                # --- Export credit warning / confirmation (V3.1 FIXED) ---
                export_cost_note = gr.Markdown("", visible=False, elem_id="export_credit_note")
                export_confirm_cb = gr.Checkbox(
                    label="I understand exporting may spend credits for image generation",
                    value=False,
                    visible=False,
                    elem_id="export_credit_confirm"
                )

            with gr.Row():
                templates_tab_btn = gr.Button("📚 Templates (Upload / Save / Load / Reference PDFs)")

            status = gr.Textbox(label="Status / Errors", lines=3)

            rate_limit_display = gr.Textbox(
                label="📊 Daily Usage",
                value="Rate limits: 50 generations/day, 100 transcriptions/day",
                interactive=False,
                lines=2
            )
            transcript_out = gr.Textbox(label="Transcript / Input used", lines=3)

            # Banner + metadata used for save
            editor_draft_banner = gr.Markdown("#### Draft: *(none loaded yet)*")

            with gr.Row():
                draft_name = gr.Textbox(label="Draft name (used on save)", placeholder="e.g. Year 11 Methods Test – Differentiation")
                draft_subject_box = gr.Textbox(label="Subject", value="", visible=False, interactive=False)

            preview_html = gr.File(label="Preview (HTML with MathJax)")

            with gr.Row():
                edited_md = gr.Textbox(label="Editable Document Markdown", lines=20)
                edited_ppt = gr.Textbox(label="Editable PPT Outline", lines=20)

            # =========================
            # Export UI State Wiring - FIXED with debouncing
            # =========================

            # When editor content changes - FULL update
            edited_md.change(
                fn=_update_export_ui,
                    inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            )

            edited_ppt.change(
                fn=_update_export_ui,
                    inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                    outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            )

            # When checkbox changes - run FULL update so buttons unlock reliably
            export_confirm_cb.change(
                fn=_update_export_ui,
                inputs=[edited_md, edited_ppt, export_snooze, export_confirm_cb, supabase_session],
                outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            ).then(
                fn=lambda v: v,
                inputs=[export_confirm_cb],
                outputs=[export_confirm_state]
            )

            with gr.Row():
                docx_file = gr.File(label="Download DOCX (editable Word equations)")
                pptx_file = gr.File(label="Download PPTX (equations rendered)")

            templates_tab_btn.click(fn=lambda: _goto_tab("Templates"), inputs=[], outputs=[tabs])

            # PDF → Draft wiring
            def _can_generate_pdf_now(pdf_file, subj, other_subj, education_level):
                if pdf_file is None:
                    return False
                if education_level == "University / Tertiary":
                    return True
                return is_subject_ready(subj, other_subj)

            course.change(
                fn=lambda f, subj, other_subj, ed: gr.update(interactive=_can_generate_pdf_now(f, subj, other_subj, ed)),
                inputs=[pdf_upload, course, other_subject, education_level],
                outputs=pdf_generate_btn
            )

            other_subject.change(
                fn=lambda f, subj, other_subj, ed: gr.update(interactive=_can_generate_pdf_now(f, subj, other_subj, ed)),
                inputs=[pdf_upload, course, other_subject, education_level],
                outputs=pdf_generate_btn
            )
            
            pdf_generate_btn.click(
                fn=action_generate_from_pdf,
                inputs=[
                    supabase_session,
                    pdf_upload,
                    pdf_draft_title,
                    education_level,
                    country, state_province,
                    uni_country, university_name, faculty, module_code,
                    year_level, course, other_subject, course_stream, output_type,
                    include_memo, model_name
                ],
                outputs=[
                    transcript_out,
                    edited_md,
                    edited_ppt,
                    preview_html,
                    status,
                    last_instruction,
                    last_mode,
                    last_transcript,
                    current_draft_id,
                    current_version
            ]).then(
                  fn=get_rate_limit_display,
                  inputs=[supabase_session],
                  outputs=[rate_limit_display]
              ).then(
                  fn=_reset_export_guard,
                  inputs=[],
                  outputs=[export_snooze, export_confirm_state]
              ).then(
                  fn=_update_export_ui,  # Use FULL update here
                      inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                  outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]

              )

            docx_generate_btn.click(
                fn=action_generate_from_docx,
                inputs=[
                    supabase_session,
                    docx_upload,
                    docx_draft_title,
                    docx_mode,
                    education_level,
                    country, state_province,
                    uni_country, university_name, faculty, module_code,
                    year_level, course, other_subject, course_stream, output_type,
                    include_memo, model_name
                ],
                outputs=[
                    transcript_out,
                    edited_md,
                    edited_ppt,
                    preview_html,
                    status,
                    last_instruction,
                    last_mode,
                    last_transcript,
                    current_draft_id,
                    current_version
            ]).then(
                  fn=get_rate_limit_display,
                  inputs=[supabase_session],
                  outputs=[rate_limit_display]
              ).then(
                  fn=_reset_export_guard,
                  inputs=[],
                  outputs=[export_snooze, export_confirm_state]
              ).then(
                  fn=_update_export_ui,  # Use FULL update here
                      inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                  outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
              )

            pptx_generate_btn.click(
                fn=action_generate_from_pptx,
                inputs=[
                    supabase_session,
                    pptx_upload,
                    pptx_draft_title,
                    education_level,
                    country, state_province,
                    uni_country, university_name, faculty, module_code,
                    year_level, course, other_subject, course_stream, output_type,
                    include_memo, model_name
                ],
                outputs=[
                    transcript_out,
                    edited_md,
                    edited_ppt,
                    preview_html,
                    status,
                    last_instruction,
                    last_mode,
                    last_transcript,
                    current_draft_id,
                    current_version
            ]).then(
                  fn=get_rate_limit_display,
                  inputs=[supabase_session],
                  outputs=[rate_limit_display]
              ).then(
                  fn=_reset_export_guard,
                  inputs=[],
                  outputs=[export_snooze, export_confirm_state]
              ).then(
                  fn=_update_export_ui,  # Use FULL update here
                      inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                  outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
              )

            # Generate action (writes into editor) — CHAINED so UI refresh always runs
            gen_btn.click(
                fn=action_generate_draft,
                inputs=[
                    supabase_session,
                    input_mode, audio_in, typed_prompt, live_transcript,
                    education_level,
                    country, state_province,
                    uni_country, university_name, faculty, module_code,
                    year_level, course, other_subject, course_stream, output_type,
                    include_memo, model_name,
                    draft_mode, current_draft_id, edited_md
                ],
                outputs=[
                    transcript_out,
                    edited_md,
                    edited_ppt,
                    preview_html,
                    status,
                    last_instruction,
                    last_mode,
                    last_transcript,
                    current_draft_id,
                    current_version
                ]
            ).then(
                fn=get_rate_limit_display,
                inputs=[supabase_session],
                outputs=[rate_limit_display]
            ).then(
                fn=_reset_export_guard,
                inputs=[],
                outputs=[export_snooze, export_confirm_state]
            ).then(
                fn=_update_export_ui,  # Use FULL update here
                    inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            )

            def _can_generate_now(m, a, t, subj, other_subj):
                return can_generate(m, a, t) and is_subject_ready(subj, other_subj)


        # =========================
        # DRAFTS TAB
        # =========================
        with gr.TabItem("Drafts", id="Drafts", visible=False) as drafts_tab:
            gr.Markdown("## 📚 Draft Library")
            drafts_lock_note = gr.Markdown("", visible=False)

            gr.Markdown("### 1) Find a draft")
            draft_search = gr.Textbox(
                label="Search drafts",
                placeholder="Type part of the draft name or subject…",
                lines=1
            )

            with gr.Row():
                refresh_btn = gr.Button("🔄 Refresh drafts")
                drafts_dd = gr.Dropdown(
                    label="My drafts",
                    choices=[],
                    interactive=True
                )

            gr.Markdown("### 2) Choose a version")
            selected_draft_banner = gr.Markdown("### 📄 Working on: *(none loaded yet)*")

            versions_dd = gr.Dropdown(
                label="Versions (newest first)",
                choices=[],
                interactive=True
            )

            load_btn = gr.Button("⬇️ Load into Workspace")

            with gr.Accordion("✏️ Rename selected draft", open=False):
                rename_new_title = gr.Textbox(
                    label="New draft name",
                    placeholder="Type the new name…",
                    lines=1
                )
                rename_btn = gr.Button("Rename draft")

            with gr.Accordion("🗑️ Delete options", open=False):
                with gr.Row():
                    delete_version_btn = gr.Button("🗑️ Delete version")
                    delete_draft_btn = gr.Button("🔥 Delete draft")

            library_status = gr.Textbox(label="Library status", lines=2)

            # Save-as-Draft from current editor
            with gr.Accordion("💾 Save current draft as reusable content", open=False):
                gr.Markdown(
                    "**Save the current draft content so you can reuse it later.**\n\n"
                    "Use this for teaching content such as tests, worksheets, assignments, memos, and exam papers.\n\n"
                    "This saves reusable **content**. It is not a school formatting template."
                )
                template_name = gr.Textbox(label="Reusable Content Name*", placeholder="e.g., Year 11 Methods Quiz")
                template_desc = gr.Textbox(label="Description (optional)", placeholder="What is this content for?")
                template_category = gr.Dropdown(
                    label="Category",
                    choices=TEMPLATE_CATEGORIES,
                    value="Test/Quiz"
                )
                share_template = gr.Checkbox(label="Share with other teachers (make public)", value=False)
                save_template_confirm = gr.Button("💾 Save Reusable Content", variant="primary")
                save_template_status = gr.Textbox(label="Save Status", interactive=False, lines=2)

            # 🔒 Plan lock wiring (Drafts tab)
            supabase_session.change(
                fn=_apply_drafts_plan_lock,
                inputs=[supabase_session],
                outputs=[
                    drafts_lock_note,
                    draft_search,
                    refresh_btn,
                    drafts_dd,
                    versions_dd,
                    load_btn,
                    rename_new_title,
                    rename_btn,
                    delete_version_btn,
                    delete_draft_btn,
                    library_status,
                    template_name,
                    template_desc,
                    template_category,
                    share_template,
                    save_template_confirm,
                    save_template_status
                ]
            )


            def ui_supa_refresh_with_search(sess, search_text):
                choices, msg = supa_list_my_drafts(sess, search_text=(search_text or ""))
                value = choices[0] if choices else None
                if not choices:
                    msg = "ℹ️ No drafts yet. Go to Workspace and generate your first document."

                return gr.update(choices=choices, value=value), msg

            refresh_btn.click(
                fn=ui_supa_refresh_with_search,
                inputs=[supabase_session, draft_search],
                outputs=[drafts_dd, library_status]
            )

            save_template_confirm.click(
                fn=save_as_template,
                inputs=[
                    supabase_session,
                    template_name,
                    template_desc,
                    template_category,
                    draft_subject_state,
                    edited_md,
                    edited_ppt,
                    share_template
                ],
                outputs=[save_template_status]
            )

            draft_search.change(
                fn=ui_supa_refresh_with_search,
                inputs=[supabase_session, draft_search],
                outputs=[drafts_dd, library_status]
            )

            drafts_dd.change(
                fn=ui_supa_versions,
                inputs=[supabase_session, drafts_dd],
                outputs=[versions_dd, library_status]
            ).then(
                fn=_banner_md_from_choice,
                inputs=[drafts_dd],
                outputs=[selected_draft_banner]
            ).then(
                fn=lambda choice: f"#### Draft: **{_draft_title_from_choice(choice) or '(none)'}**",
                inputs=[drafts_dd],
                outputs=[editor_draft_banner]
            ).then(
                fn=lambda choice: gr.update(value=_draft_title_from_choice(choice) or ""),
                inputs=[drafts_dd],
                outputs=[draft_name]
            ).then(
                fn=_set_subject_state_from_choice,
                inputs=[drafts_dd],
                outputs=[draft_subject_state, draft_subject_box]
            )

            def _load_selected_version_and_go(sess, draft_choice, version_choice):
                md, ppt, msg, draft_id, version_num = supa_load_selected_version(sess, draft_choice, version_choice)
                return md, ppt, msg, draft_id, version_num, "Workspace"


            load_btn.click(
                fn=_load_selected_version_and_go,
                inputs=[supabase_session, drafts_dd, versions_dd],
                outputs=[edited_md, edited_ppt, library_status, current_draft_id, current_version, current_tab]
            ).then(
                fn=lambda tab: gr.update(selected=tab),
                inputs=[current_tab],
                outputs=[tabs]
            ).then(
                fn=_banner_md_from_choice,
                inputs=[drafts_dd],
                outputs=[selected_draft_banner]
            ).then(
                fn=lambda choice: f"#### Draft: **{_draft_title_from_choice(choice) or '(none)'}**",
                inputs=[drafts_dd],
                outputs=[editor_draft_banner]
            ).then(
                fn=lambda choice: gr.update(value=_draft_title_from_choice(choice) or ""),
                inputs=[drafts_dd],
                outputs=[draft_name]
            ).then(
                fn=_reset_export_guard,
                inputs=[],
                outputs=[export_snooze, export_confirm_state]
            ).then(
                fn=_update_export_ui,
                inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            )

            rename_btn.click(
                fn=supa_rename_draft,
                inputs=[supabase_session, drafts_dd, rename_new_title],
                outputs=[library_status]
            ).then(
                fn=ui_supa_refresh_with_search,
                inputs=[supabase_session, draft_search],
                outputs=[drafts_dd, library_status]
            )

            delete_version_btn.click(
                fn=supa_delete_version,
                inputs=[supabase_session, drafts_dd, versions_dd],
                outputs=[library_status]
            ).then(
                fn=ui_supa_versions,
                inputs=[supabase_session, drafts_dd],
                outputs=[versions_dd, library_status]
            )

            delete_draft_btn.click(
                fn=supa_delete_draft,
                inputs=[supabase_session, drafts_dd],
                outputs=[library_status]
            ).then(
                fn=ui_supa_refresh_with_search,
                inputs=[supabase_session, draft_search],
                outputs=[drafts_dd, library_status]
            ).then(
                fn=lambda: ("", "", "", 0),
                inputs=[],
                outputs=[edited_md, edited_ppt, current_draft_id, current_version]
            ).then(
                fn=lambda: "### 📄 Selected draft: *(none loaded yet)*",
                inputs=[],
                outputs=[selected_draft_banner]
            ).then(
                fn=lambda: "#### Draft: *(none loaded yet)*",
                inputs=[],
                outputs=[editor_draft_banner]
            ).then(
                fn=lambda: gr.update(value=""),
                inputs=[],
                outputs=[draft_name]
            )

        # ======================================
        # TEMPLATES TAB (moved out of Workspace)
        # ======================================
        with gr.TabItem("Templates", id="Templates", visible=False) as templates_tab:
            gr.Markdown("## 📚 Templates")
            templates_lock_note = gr.Markdown("", visible=False)

            # Upload scaffold
            with gr.Accordion("⬆️ Upload your school document format", open=False):
                gr.Markdown(
                    "**Upload a document that represents your school's format.**\n\n"
                    "This could be an exam paper, worksheet, or test that already follows your school's layout.\n\n"
                    "The system will use this as a reference structure so future drafts can follow a similar format.\n\n"
                    "**Note:** This currently captures structure and content layout. Full visual formatting (like exact spacing and styling from PDFs) is not yet applied."
                )
                up_template_name = gr.Textbox(label="Template Name*", placeholder="e.g., Year 8 Exam Scaffold - Calc Free")
                up_template_desc = gr.Textbox(label="Description (optional)", placeholder="e.g., Uses our Semester 2 exam structure with marks format")
                up_template_category = gr.Dropdown(
                    label="Category",
                    choices=TEMPLATE_CATEGORIES,
                    value="Custom"
                )
                up_share_template = gr.Checkbox(label="Share with other teachers (make public)", value=False)

                up_file = gr.File(
                    label="Upload school document (DOCX, PDF, or PPT)",
                    file_types=[".docx", ".pdf", ".pptx", ".ppt"],
                    type="filepath"
                )

                upload_template_confirm = gr.Button("⬆️ Extract School Template", variant="primary")
                save_analyzed_template_btn = gr.Button("💾 Save Analyzed Template to Library", variant="secondary")
                upload_template_status = gr.Textbox(label="Upload Status", interactive=False, lines=2)
                
                template_clean_md_state = gr.State("")
                template_profile_json_state = gr.State("")
                template_bundle_state = gr.State("")

                template_clean_preview = gr.Textbox(
                    label="Extracted Institution Style Preview",
                    lines=14,
                    interactive=False
                )

                template_profile_preview = gr.Textbox(
                    label="Extracted Template Profile (JSON)",
                    lines=18,
                    interactive=False
                )                

            # Load / Delete templates
            with gr.Accordion("📂 Manage Templates", open=False):
                template_category_filter = gr.Dropdown(
                    label="Filter by Category",
                    choices=["All"] + TEMPLATE_CATEGORIES,
                    value="All"
                )
                templates_dropdown = gr.Dropdown(label="My Templates", choices=[], interactive=True)
                refresh_templates_btn = gr.Button("🔄 Remember to Refresh your Templates")
                apply_template_btn = gr.Button("✨ Apply Template to Current Draft", variant="primary")
                delete_template_btn = gr.Button("🗑️ Delete Template", variant="stop")
                load_template_status = gr.Textbox(label="Template Status", interactive=False, lines=2)
                generated_template_file = gr.File(label="Generated Template Output", interactive=False, visible=True)

            # ============================================================
            # TEACHER CONFIRMATION PANEL (for uncertain decisions)
            # ============================================================
            with gr.Group(visible=False) as confirmation_panel:
                gr.Markdown("## ⚠️ Teacher Confirmation Required")
                gr.Markdown("The engine is uncertain about some layout decisions. Please review below:")
                confirmation_prompt_display = gr.Textbox(
                    label="Engine's Analysis", 
                    lines=12, 
                    interactive=False,
                    placeholder="Confirmation prompt will appear here..."
                )
                confirmation_choice = gr.Radio(
                    choices=["A", "B", "C", "D"],
                    label="Your decision",
                    info="A: Accept engine's recommendation | B: Apply this feature | C: Do not apply this feature | D: Keep donor's original layout"
                )
                confirm_submit_btn = gr.Button("✅ Submit & Continue", variant="primary")
                confirm_status = gr.Textbox(label="Status", lines=2, interactive=False)

            # Reference PDF attach/load
            with gr.Accordion("📎 Attach Reference PDF to a Template (optional)", open=False):
                gr.Markdown("**Attach a PDF layout reference to an existing template (PDF is NOT parsed)**")

                ref_template_dropdown = gr.Dropdown(
                    label="Select template to attach PDF to",
                    choices=[],
                    interactive=True
                )

                ref_refresh_templates_btn = gr.Button("🔄 Refresh templates (for PDF attach)")

                ref_pdf_upload_attach = gr.File(
                    label="Upload reference PDF (.pdf only)",
                    file_types=[".pdf"]
                )

                ref_attach_btn = gr.Button("📎 Attach PDF", variant="primary")
                ref_load_btn = gr.Button("⬇️ Load attached PDF", variant="secondary")

                ref_pdf_file = gr.File(label="Attached Reference PDF (download/open)")
                ref_pdf_status = gr.Textbox(label="Reference PDF Status", interactive=False, lines=2)

            # Template wiring
            refresh_templates_btn.click(
                fn=refresh_templates_on_open,
                inputs=[supabase_session, template_category_filter],
                outputs=[templates_dropdown, load_template_status]
            )

            template_category_filter.change(
                fn=refresh_templates_on_open,
                inputs=[supabase_session, template_category_filter],
                outputs=[templates_dropdown, load_template_status]
            )

            def _analyze_template_bridge(uploaded_file, model_name_value):
                status, clean_md, profile_json, bundle, _save_btn_update = action_analyze_template_upload(
                    uploaded_file,
                    model_name_value
                )
                return status, clean_md, profile_json, clean_md, profile_json, bundle


            upload_template_confirm.click(
                fn=_analyze_template_bridge,
                inputs=[up_file, model_name],
                outputs=[
                    upload_template_status,
                    template_clean_preview,
                    template_profile_preview,
                    template_clean_md_state,
                    template_profile_json_state,
                    template_bundle_state,
                ]
            )

            def _save_analyzed_template_bridge(sess, name, desc, category, bundle, file_obj, cat_filter):
                status = save_template_record(
                    sess,
                    name,
                    desc,
                    category,
                    bundle,
                    file_obj
                )
                dd_update, load_msg = refresh_templates_on_open(sess, cat_filter)
                return status, dd_update, load_msg


            save_analyzed_template_btn.click(
                fn=_save_analyzed_template_bridge,
                inputs=[
                    supabase_session,
                    up_template_name,
                    up_template_desc,
                    up_template_category,
                    template_bundle_state,
                    up_file,
                    template_category_filter,
                ],
                outputs=[
                    upload_template_status,
                    templates_dropdown,
                    load_template_status,
                ]
            )

            def _load_template_into_preview(template_choice):
                if isinstance(template_choice, list):
                    template_choice = template_choice[0] if template_choice else ""

                if not template_choice:
                    return "", "", "❌ No template selected.", "", "", ""

                if "|" in template_choice:
                    template_id = template_choice.split("|", 1)[1].strip()
                else:
                    template_id = template_choice.strip()

                clean_md, profile, status = load_template_bundle_from_db(template_id)

                import json

                pretty_profile = json.dumps(profile, ensure_ascii=False, indent=2) if profile else ""

                # IMPORTANT:
                # - DOCX-native templates must keep their metadata JSON exactly as-is
                # - legacy templates can still use the old packed markdown bundle
                if isinstance(profile, dict) and profile.get("storage_mode") == "docx_native_v1":
                    bundle = json.dumps({
                        "storage_mode": "docx_native_v1",
                        "source_type": profile.get("source_type", "docx"),
                        "source_filename": profile.get("source_filename", ""),
                        "template_title": profile.get("template_title", ""),
                        "description": profile.get("description", ""),
                        "category": profile.get("category", ""),
                        "docx_storage_path": profile.get("docx_storage_path", ""),
                        "docx_signed_url_at_save": profile.get("docx_signed_url_at_save", ""),
                        "saved_at": profile.get("saved_at", ""),
                        "profile": profile,
                        "legacy_clean_md_preview": clean_md or "",
                    }, ensure_ascii=False, indent=2)
                else:
                    bundle = pack_template_bundle(clean_md, profile) if profile else clean_md

                return clean_md, pretty_profile, status, clean_md, pretty_profile, bundle

            templates_dropdown.change(
                fn=_load_template_into_preview,
                inputs=[templates_dropdown],
                outputs=[
                    template_clean_preview,
                    template_profile_preview,
                    load_template_status,
                    template_clean_md_state,
                    template_profile_json_state,
                    template_bundle_state,
                ]
            )

            bs5_debug_state = gr.State("")

            def _apply_template_bridge(current_draft_md, template_bundle, model_name_value, year_level_value, subject_value):
                try:
                    from template_engine import debug_template_engine_snapshot
                    clean_year_level = (year_level_value or "").strip()
                    clean_subject = (subject_value or "").strip()

                    if clean_year_level == "Year 11":
                        clean_year_level = ""

                    debug_snapshot = debug_template_engine_snapshot(
                        current_draft_md=current_draft_md,
                        template_bundle=template_bundle,
                        year_level=clean_year_level,
                        subject=clean_subject
                    )
                except Exception as e:
                    debug_snapshot = (
                        '{\n'
                        f'  "status": "error",\n'
                        f'  "error_type": "{type(e).__name__}",\n'
                        f'  "error_message": "{str(e)}"\n'
                        '}'
                    )

                clean_year_level = (year_level_value or "").strip()
                clean_subject = (subject_value or "").strip()

                # BS5 RULE:
                # Do not inject stale UI defaults into template_engine.
                # Let the draft stay the content authority unless the current
                # values are genuinely present and intentional.
                if clean_year_level == "Year 11":
                    clean_year_level = ""

                result = apply_template_to_draft(
                    current_draft_md=current_draft_md,
                    template_bundle=template_bundle,
                    model_name=model_name_value,
                    year_level=clean_year_level,
                    subject=clean_subject
                )

                # Check if result is a confirmation prompt (starts with uncertainty indicator)
                if isinstance(result, str) and result.startswith("The engine is uncertain"):
                    return (
                        current_draft_md,           # editor unchanged
                        result,                     # show prompt in status
                        gr.update(value=None, visible=False),  # hide file output
                        gr.update(visible=True),    # show confirmation panel
                        result,                     # prompt in display box
                        "",                         # clear confirm status
                        debug_snapshot              # BS5 debug snapshot
                    )

                # DOCX result: real file path
                if isinstance(result, str) and result.lower().endswith(".docx") and os.path.exists(result):
                    return (
                        current_draft_md,
                        "✅ DOCX template applied successfully.",
                        gr.update(value=result, visible=True),
                        gr.update(visible=False),   # hide confirmation panel
                        "",
                        "",
                        debug_snapshot
                    )

                # Error result
                if isinstance(result, str) and result.startswith(("❌", "⚠️")):
                    return (
                        current_draft_md,
                        result,
                        None,
                        gr.update(visible=False),
                        "",
                        "",
                        debug_snapshot
                    )

                # Markdown fallback
                return (
                    result,
                    "✅ Template applied using blueprint mode.",
                    None,
                    gr.update(visible=False),
                    "",
                    "",
                    debug_snapshot
                )

            # Update the click handler to use the new outputs
            apply_template_btn.click(
                fn=_apply_template_bridge,
                inputs=[edited_md, template_bundle_state, model_name, shared_year_level, shared_subject],
                outputs=[
                    edited_md, 
                    load_template_status, 
                    generated_template_file,
                    confirmation_panel,
                    confirmation_prompt_display,
                    confirm_status,
                    bs5_debug_state
                ]
            ).then(
                fn=_reset_export_guard,
                inputs=[],
                outputs=[export_snooze, export_confirm_state]
            ).then(
                fn=_update_export_ui,
                inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            )

            # Handle confirmation submission
            def _handle_confirmation_submit(choice, prompt_text, draft_md, template_bundle):
                if not choice:
                    return "⚠️ Please select an option (A, B, C, or D).", None, gr.update(visible=True)

                # For the current BS5 live flow, the confirmation UI is handling
                # a single explicit uncertainty surfaced by the engine.
                # Build the minimal uncertain_decisions payload expected by
                # handle_teacher_confirmation(...).
                uncertain_decisions = {
                    "textboxes": Decision(
                        apply=True,
                        confidence=0.50,
                        reason="Contextual feature 'textboxes' has no specific Phase 2 rule yet; defaulting to apply.",
                        alternative="Confirm whether 'textboxes' should be applied.",
                        requires_teacher_confirmation=True,
                        source="engine",
                    )
                }

                try:
                    result = handle_teacher_confirmation(
                        confirmation_response=choice,
                        uncertain_decisions=uncertain_decisions,
                        draft_md=draft_md,
                        template_bundle=template_bundle
                    )
                except Exception as e:
                    return (
                        f"❌ Confirmation submit failed: {type(e).__name__}: {e}",
                        None,
                        gr.update(visible=True)
                    )

                if isinstance(result, str) and result.lower().endswith(".docx") and os.path.exists(result):
                    return (
                        "✅ Document generated with your preferences.",
                        gr.update(value=result, visible=True),
                        gr.update(visible=False)
                    )

                return (
                    f"⚠️ Failed: {result}",
                    None,
                    gr.update(visible=True)
                )
            
            confirm_submit_btn.click(
                fn=_handle_confirmation_submit,
                inputs=[confirmation_choice, confirmation_prompt_display, edited_md, template_bundle_state],
                outputs=[confirm_status, generated_template_file, confirmation_panel]
            )

            delete_template_btn.click(
                fn=delete_template,
                inputs=[supabase_session, templates_dropdown],
                outputs=[load_template_status]
            ).then(
                fn=refresh_templates_on_open,
                inputs=[supabase_session, template_category_filter],
                outputs=[templates_dropdown, load_template_status]
            )

            ref_refresh_templates_btn.click(
                fn=refresh_templates_on_open,
                inputs=[supabase_session, template_category_filter],
                outputs=[ref_template_dropdown, load_template_status]
            )

            ref_attach_btn.click(
                fn=attach_reference_pdf_to_template,
                inputs=[supabase_session, ref_template_dropdown, ref_pdf_upload_attach],
                outputs=[ref_pdf_status, ref_pdf_file]
            )
            ref_load_btn.click(
                fn=load_reference_pdf_for_template,
                inputs=[supabase_session, ref_template_dropdown],
                outputs=[ref_pdf_status, ref_pdf_file]
            )

            # 🔒 Plan lock wiring (Templates tab)
            supabase_session.change(
                fn=_apply_templates_plan_lock,
                inputs=[supabase_session],
                outputs=[
                    templates_lock_note,
                    upload_template_confirm,
                    save_analyzed_template_btn,
                    refresh_templates_btn,
                    apply_template_btn,
                    delete_template_btn,
                    ref_attach_btn,
                    ref_load_btn,
                    templates_dropdown,
                    ref_template_dropdown,
                    up_file,
                    load_template_status,
                    generated_template_file
                ]
            )

        from datetime import datetime, timezone

        def _utc_now_iso():
            return datetime.now(timezone.utc).isoformat()

        def _uid_from_session(sess):
            """
            Tries to extract user_id from the session object you already pass around (supabase_session).
            Handles common shapes safely.
            """
            if not sess:
                return None

            # Common: {"user": {"id": "..."}}
            if isinstance(sess, dict):
                u = sess.get("user")
                if isinstance(u, dict) and u.get("id"):
                    return u.get("id")

                # Common: {"user_id": "..."} or {"id": "..."}
                if sess.get("user_id"):
                    return sess.get("user_id")
                if sess.get("id"):
                    return sess.get("id")

                # Sometimes nested: {"session": {"user": {"id": ...}, "access_token": ...}}
                s2 = sess.get("session")
                if isinstance(s2, dict):
                    u2 = s2.get("user")
                    if isinstance(u2, dict) and u2.get("id"):
                        return u2.get("id")

            return None

        def _access_token_from_session(sess):
            if not sess or not isinstance(sess, dict):
                return None

            # Common token locations
            if sess.get("access_token"):
                return sess.get("access_token")

            s2 = sess.get("session")
            if isinstance(s2, dict) and s2.get("access_token"):
                return s2.get("access_token")

            return None


        def _refresh_token_from_session(sess):
            if not sess or not isinstance(sess, dict):
                return None

            # Common token locations
            if sess.get("refresh_token"):
                return sess.get("refresh_token")

            s2 = sess.get("session")
            if isinstance(s2, dict) and s2.get("refresh_token"):
                return s2.get("refresh_token")

            return None


        def _sb_authed_from_session(sess):
            """
            Creates a Supabase client and attaches the user's session so both
            PostgREST and Storage requests run under the user's JWT/RLS context.
            """
            from supabase import create_client
            import os

            url = (
                globals().get("SUPABASE_URL")
                or globals().get("SUPABASE_PROJECT_URL")
                or os.getenv("SUPABASE_URL")
                or os.getenv("SUPABASE_PROJECT_URL")
                or ""
            )

            key = (
                globals().get("SUPABASE_ANON_KEY")
                or globals().get("SUPABASE_KEY")
                or globals().get("SUPABASE_ANON")
                or os.getenv("SUPABASE_ANON_KEY")
                or os.getenv("SUPABASE_KEY")
                or os.getenv("SUPABASE_ANON")
                or ""
            )

            url = (url or "").strip()
            key = (key or "").strip()

            if not url or not key:
                raise NameError(
                    "Supabase config missing. Set SUPABASE_URL + SUPABASE_ANON_KEY in Secrets/Env (or your app's equivalent)."
                )

            sb = create_client(url, key)

            access_token = _access_token_from_session(sess)
            refresh_token = _refresh_token_from_session(sess)

            # Best path: attach the full session so Storage and DB both use user auth
            if access_token and refresh_token:
                try:
                    sb.auth.set_session(access_token, refresh_token)
                except Exception:
                    # Fallback: at least keep PostgREST authenticated
                    sb.postgrest.auth(access_token)
            elif access_token:
                sb.postgrest.auth(access_token)

            return sb

        # =========================
        # LEGAL ACKNOWLEDGEMENTS (Profile tab)
        # =========================

        LEGAL_DOCS = {
            # doc_key: (label, profiles_ack_at_col, profiles_version_col, hardcoded_version_string, legal_ack_events.doc_type)
            "privacy": ("Privacy Policy", "privacy_policy_ack_at", "privacy_policy_version", LEGAL_VERSIONS["privacy"], "privacy_policy"),
            "terms":   ("Terms of Use", "terms_of_use_ack_at", "terms_of_use_version", LEGAL_VERSIONS["terms"], "terms_of_use"),
            "tcs":     ("Terms & Conditions", "terms_and_conditions_ack_at", "terms_and_conditions_version", LEGAL_VERSIONS["tcs"], "terms_and_conditions"),
        }


        def _fmt_ack_legal_event(ts_iso: str, ver: str) -> str:
            ts_iso = (ts_iso or "").strip()
            ver = (ver or "").strip()
            if not ts_iso:
                return "❌ Not acknowledged yet."
            # Keep it simple and robust (no timezone parsing dependencies)
            return f"✅ Acknowledged on **{ts_iso}** (version **{ver or 'unknown'}**)."

        def action_ack_legal_doc(sess, doc_key: str, checked: bool):
            """
            When user ticks a legal checkbox:
              - profiles.<doc>_ack_at = now (UTC iso)
              - profiles.<doc>_version = hardcoded version string
              - (optional) insert into legal_ack_events if table exists
            If user unticks, we do NOTHING (we don't allow "un-ack").
            Returns: (status_md, checkbox_update)
            """
            if not sess:
                return "❌ Not signed in.", gr.update(value=False, interactive=False)

            if not checked:
                # Do not clear acknowledgement once set
                return "ℹ️ Acknowledgement cannot be removed once recorded.", gr.update(value=True, interactive=False)

            doc_key = (doc_key or "").strip().lower()
            if doc_key not in LEGAL_DOCS:
                return "❌ Unknown legal document key.", gr.update(value=False)

            label, ack_col, ver_col, ver_fallback, doc_type = LEGAL_DOCS[doc_key]

            # Source of truth for required version = DB table: legal_config
            # Fallback = LEGAL_DOCS hardcoded version (ver_fallback)
            ver = None
            try:
                sb_tmp = _sb_authed_from_session(sess)
                row = (
                    sb_tmp.table("legal_config")
                    .select("current_version")
                    .eq("doc_type", doc_key)
                    .limit(1)
                    .execute()
                )
                if row.data and row.data[0].get("current_version"):
                    ver = (row.data[0].get("current_version") or "").strip() or None
            except Exception:
                ver = None

            ver = ver or ver_fallback

            uid = _uid_from_session(sess)
            if not uid:
                return "❌ Could not determine your user id from session.", gr.update(value=False)

            try:
                sb = _sb_authed_from_session(sess)
                now_iso = _utc_now_iso()

                # Update profile
                sb.table("profiles").update({
                    ack_col: now_iso,
                    ver_col: ver
                }).eq("user_id", uid).execute()

                # Optional audit trail (only if the table exists)
                try:
                    sb.table("legal_ack_events").insert({
                        "user_id": uid,
                        "doc_type": doc_type,
                        "doc_version": ver,
                        "ack_at": now_iso,
                        "client_meta": {"surface": "profile_tab"}
                    }).execute()
                except Exception:
                    pass

                return _fmt_ack_legal_event(now_iso, ver), gr.update(value=True, interactive=False)

            except Exception as e:
                # Don't flip the checkbox back on failure (that can cause UI churn/loops)
                return f"❌ Failed to save acknowledgement: {e}", gr.update(value=True, interactive=True)

        def action_activate_pro(sess):
            """
            Temporary internal Pro activation.
            Updates the user's profiles row so the existing plan locks,
            badges, and access checks immediately treat the user as Pro.

            Later, this can be replaced by Stripe/payment webhook logic.
            """
            if not sess:
                return "❌ Not signed in."

            uid = _uid_from_session(sess)
            if not uid:
                return "❌ Could not determine your user id from session."

            try:
                # Make sure the profile row exists first
                ensure_profile_row(sess)

                sb = _sb_authed_from_session(sess)
                now = datetime.now(timezone.utc).isoformat()

                sb.table("profiles").update({
                    "is_pro": True,
                    "plan": "pro",
                    "subscription_plan": "pro",
                    "updated_at": now
                }).eq("user_id", uid).execute()

                return "✅ Pro activated successfully."

            except Exception as e:
                return f"❌ Pro activation failed: {e}"

        def action_deactivate_pro(sess):
            """
            Self-service downgrade to Free.
            Mirrors action_activate_pro, but restores Free access state.
            """
            if not sess:
                return "❌ Not signed in."

            uid = _uid_from_session(sess)
            if not uid:
                return "❌ Could not determine your user id from session."

            try:
                ensure_profile_row(sess)
                sb = _sb_authed_from_session(sess)
                now = datetime.now(timezone.utc).isoformat()

                sb.table("profiles").update({
                    "is_pro": False,
                    "plan": "free",
                    "subscription_plan": "free",
                    "updated_at": now
                }).eq("user_id", uid).execute()

                return "✅ Downgraded to Free successfully."

            except Exception as e:
                return f"❌ Downgrade failed: {e}"

        def action_pause_account(sess, pause_confirmed: bool):
            """
            Backend: inserts a pause request + flips profiles.is_paused true.
            (Actual auth disabling can be done later via Edge Function/admin.)
            """
            if not sess:
                return "❌ Not signed in."
            if not pause_confirmed:
                return "⚠️ Please tick the confirmation box first."

            uid = _uid_from_session(sess)
            if not uid:
                return "❌ Could not determine your user id from session."

            try:
                sb = _sb_authed_from_session(sess)

                # If already paused, just report
                prof = sb.table("profiles").select("is_paused").eq("user_id", uid).limit(1).execute()
                if prof.data and prof.data[0].get("is_paused") is True:
                    return "ℹ️ Your account is already paused."

                # 1) Insert request (audit trail / queue)
                sb.table("account_action_requests").insert({
                    "user_id": uid,
                    "action": "pause",
                    "status": "requested",
                    "client_meta": {"surface": "profile_tab"}
                }).execute()

                # 2) Update profile flags
                sb.table("profiles").update({
                    "is_paused": True,
                    "paused_at": _utc_now_iso(),
                    "pause_reason": "User requested pause via Profile tab"
                }).eq("user_id", uid).execute()

                return "✅ Pause requested. Your account is now marked as paused (you can unpause later when we enable it)."

            except Exception as e:
                return f"❌ Pause failed: {e}"


        def action_unpause_account(sess, unpause_confirmed: bool):
            """
            Backend: inserts an unpause request + flips profiles.is_paused false.
            """
            if not sess:
                return "❌ Not signed in."
            if not unpause_confirmed:
                return "⚠️ Please tick the confirmation box first."

            uid = _uid_from_session(sess)
            if not uid:
                return "❌ Could not determine your user id from session."

            try:
                sb = _sb_authed_from_session(sess)

                # If already unpaused, just report
                prof = sb.table("profiles").select("is_paused").eq("user_id", uid).limit(1).execute()
                if prof.data and prof.data[0].get("is_paused") is False:
                    return "ℹ️ Your account is already active."

                # 1) Insert request (audit trail / queue)
                sb.table("account_action_requests").insert({
                    "user_id": uid,
                    "action": "unpause",
                    "status": "requested",
                    "client_meta": {"surface": "profile_tab"}
                }).execute()

                # 2) Update profile flags
                sb.table("profiles").update({
                    "is_paused": False,
                    "paused_at": None,
                    "pause_reason": None
                }).eq("user_id", uid).execute()

                return "✅ Account reactivated. Your access is restored."

            except Exception as e:
                return f"❌ Unpause failed: {e}"


        def action_delete_account_request(sess, c1: bool, c2: bool, typed: str):
            """
            Backend: inserts a delete request + marks profiles.deletion_status = requested.
            Actual hard delete/cleanup should be done later via Edge Function/admin job.
            """
            if not sess:
                return "❌ Not signed in."
            if not (c1 and c2):
                return "⚠️ Please tick both confirmation boxes first."
            if (typed or "").strip().upper() != "DELETE":
                return "⚠️ Please type DELETE to confirm."

            uid = _uid_from_session(sess)
            if not uid:
                return "❌ Could not determine your user id from session."

            try:
                sb = _sb_authed_from_session(sess)

                # Prevent duplicate requests
                prof = sb.table("profiles").select("deletion_status").eq("user_id", uid).limit(1).execute()
                if prof.data:
                    st = (prof.data[0].get("deletion_status") or "none").lower()
                    if st in ("requested", "processing", "deleted"):
                        return f"ℹ️ Delete already {st}. No new request created."

                # 1) Insert request
                sb.table("account_action_requests").insert({
                    "user_id": uid,
                    "action": "delete",
                    "status": "requested",
                    "client_meta": {"surface": "profile_tab"}
                }).execute()

                # 2) Mark profile
                sb.table("profiles").update({
                    "deletion_status": "requested",
                    "deletion_requested_at": _utc_now_iso()
                }).eq("user_id", uid).execute()

                return "✅ Delete requested. We’ll process this when deletion is enabled (and you may be contacted if required)."

            except Exception as e:
                return f"❌ Delete request failed: {e}"


        # =========================
        # LEGAL DOC TEXT LOADER (B2)
        # =========================
        def _read_text_file(path: str) -> str:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return f.read().strip()
            except Exception as e:
                return (
                    f"*(Legal document not found at `{path}`.)*\n\n"
                    "Add the file and restart the app.\n\n"
                    f"Error: {e}"
                )

        # =========================
        # PROFILE TAB
        # =========================
        with gr.TabItem("Profile", id="Profile", visible=False) as profile_tab:
            gr.Markdown("## 👤 Profile")

            # --- Profile banner (Facebook-style) ---
            with gr.Row(elem_id="profile_banner_row"):
                banner_avatar = gr.Image(label="", height=72, width=72, elem_id="profile_banner_avatar")
                with gr.Column():
                    banner_name = gr.Markdown("**(not loaded)**")
                    banner_signed = gr.Markdown("_Not signed in._")
                    plan_badge = gr.Markdown("🆓 Free")

            profile_logout_btn = gr.Button("Log out")

            # Hidden system fields (still updated in the background by snapshot wiring)
            with gr.Group(visible=False):
                profile_header = gr.Textbox(label="Signed in as", lines=1, interactive=False)
                snap_status = gr.Textbox(label="Status", lines=1, interactive=False)

            gr.Markdown("## 🧾 Account Details")

            with gr.Accordion(" ", open=False):

                # =========================

                # Core Identity (UI-only for now)

                # =========================

                with gr.Row():

                    first_name_in = gr.Textbox(label="First name", placeholder="e.g., Jane")

                    last_name_in  = gr.Textbox(label="Last name", placeholder="e.g., Doe")


                email_in = gr.Textbox(label="Email", interactive=False, placeholder="(auto from login)")


                with gr.Row():

                    country_in = gr.Dropdown(label="Country", choices=COUNTRIES, value=None)

                    LANGUAGES = [
                        "English",
                        "Afrikaans",
                        "French",
                        "Spanish",
                        "German",
                        "Dutch",
                        "Portuguese",
                        "Italian",
                        "Greek",
                        "Arabic",
                        "Hindi",
                        "Mandarin (Chinese)",
                        "Cantonese (Chinese)",
                        "Japanese",
                        "Korean",
                        "Vietnamese",
                        "Thai",
                        "Indonesian",
                        "Malay",
                        "Filipino (Tagalog)",
                        "Turkish",
                        "Russian",
                        "Ukrainian",
                        "Polish",
                        "Other",
                    ]

                    language_in = gr.Dropdown(
                        label="Language (type to search)",
                        choices=LANGUAGES,
                        value="English",
                        filterable=True,
                        allow_custom_value=True,  # lets a teacher type a language not in the list
                    )


                years_taught_in = gr.Dropdown(

                    label="Year(s) taught",

                    choices=[f"Year {i}" for i in range(1,13)] + ["Kindergarten","Prep","University/Tertiary"],

                    multiselect=True

                )


                # =========================

                # Teaching Context (UI-only for now)

                # =========================

                with gr.Row():

                    school_type_in = gr.Dropdown(

                        label="School type",

                        choices=["Primary","Secondary","Combined","Tertiary","Other"],

                        value=None

                    )

                    curriculum_system_in = gr.Dropdown(

                        label="Curriculum system",

                        choices=["Australian Curriculum","Cambridge","IB","GCSE / A-Level","CAPS (South Africa)","NCEA","Other"],

                        value=None

                    )


                subjects_taught_in = gr.Dropdown(

                    label="Subject(s) taught",

                    choices=["Mathematics","English","Science","Biology","Physics","Chemistry","History","Geography","Economics","Other"],

                    multiselect=True

                )


                # =========================

                # Assessment Preferences (UI-only for now)

                # =========================

                assessment_style_in = gr.CheckboxGroup(

                    label="Assessment style preference",

                    choices=["Tests / Exams","Worksheets","Investigations","Projects","Homework","Revision Packs"],

                    value=[]

                )


                difficulty_pref_in = gr.Radio(

                    label="Difficulty preference",

                    choices=["Below level","At level","Mixed ability","Extension / ATAR / Advanced"],

                    value="At level"

                )


                # =========================

                # Language & Formatting (UI-only for now)

                # =========================

                with gr.Row():

                    spelling_pref_in = gr.Dropdown(

                        label="Spelling preference",

                        choices=["Australian English","British English","American English"],

                        value="Australian English"

                    )

                    marking_style_in = gr.Dropdown(

                        label="Marking / Scoring style",

                        choices=["Marks","Points","Rubric-based"],

                        value="Marks"

                    )


                preferred_tone_in = gr.Dropdown(

                    label="Preferred tone",

                    choices=["Formal","Teacher-friendly","Student-friendly"],

                    value="Teacher-friendly"

                )


                # =========================

                # Class Context (Optional)

                # =========================

                with gr.Row():

                    class_size_in = gr.Dropdown(label="Typical class size", choices=["<20","20–25","25–30","30+"], value=None)

                    ability_mix_in = gr.Dropdown(label="Student ability mix", choices=["Mostly struggling","Mixed","Mostly strong","Extension-focused"], value=None)


                # =========================

                # Compliance & Safeguards (light)

                # =========================

                confirm_educator_in = gr.Checkbox(label="I confirm I am an educator or education professional", value=False)

                confirm_review_in   = gr.Checkbox(label="I understand generated materials must be reviewed before classroom use", value=False)

                display_name_in = gr.Textbox(label="Display name", placeholder="e.g., Jane Doe")

                with gr.Row():
                    notify_weekly = gr.Checkbox(label="Weekly summary email", value=False)
                    notify_export_done = gr.Checkbox(label="Notify when export completes", value=True)
                    notify_low_credits = gr.Checkbox(label="Low credits warning", value=True)

                with gr.Row():
                    snap_btn = gr.Button("Refresh profile dashboard")
                    save_profile_btn = gr.Button("Save profile settings")

                save_profile_status = gr.Textbox(label="Save result", lines=1, interactive=False)
                saved_profile_summary = gr.Markdown("**Saved profile summary**\n\n_Not loaded yet._")

            # =========================
            # AVATAR (moved out of Account Details)
            # =========================
            gr.Markdown("## 🖼️ Avatar")

            with gr.Accordion(" ", open=False):
                avatar_preview = gr.Image(label="Current avatar", height=140)
                avatar_upload = gr.File(label="Upload new avatar (PNG/JPG/WEBP)")
                avatar_save_btn = gr.Button("Save avatar")
                avatar_status = gr.Textbox(label="Avatar status", lines=2, interactive=False)
                avatar_reload_btn = gr.Button("Reload avatar")
            
            gr.Markdown("## 📚 Your Saved Work")

            # Keep count for backend wiring (snapshot outputs), but hide it.
            drafts_count_out = gr.Textbox(label="Drafts count", lines=1, interactive=False, visible=False)

            # Visible UI: this will become the accordion display after snapshot runs
            drafts_list_out = gr.Markdown(value="(No drafts yet)")


            with gr.Accordion("🧩 Your Templates", open=False):

                # Keep count for backend wiring (snapshot outputs), but hide it.
                templates_count_out = gr.Textbox(
                    label="Templates count",
                    lines=1,
                    interactive=False,
                    visible=False
                )

                # Visible UI: this will become the accordion display after snapshot runs
                templates_list_out = gr.Markdown(value="(No templates yet)")

                gr.Markdown("### 🔎 Template Viewer (Preview + Load)")

                with gr.Row():
                    profile_template_category_filter = gr.Dropdown(
                        label="Filter by Category",
                        choices=["All"] + TEMPLATE_CATEGORIES,
                        value="All"
                    )
                    profile_templates_dd = gr.Dropdown(
                        label="My Templates",
                        choices=[],
                        interactive=True
                    )

                with gr.Row():
                    profile_templates_refresh_btn = gr.Button("🔄 Refresh templates")
                    profile_preview_template_btn = gr.Button("👁 Preview template")
                    profile_load_template_btn = gr.Button("⬇️ Load into Workspace Editor", variant="primary")

                profile_template_status = gr.Textbox(
                    label="Template viewer status",
                    lines=2,
                    interactive=False
                )

                with gr.Row():
                    profile_template_md = gr.Textbox(
                        label="Template Document (read-only)",
                        lines=14,
                        interactive=False
                    )
                    profile_template_ppt = gr.Textbox(
                        label="Template PPT Outline (read-only)",
                        lines=14,
                        interactive=False
                    )


            # =========================
            # ACCOUNT SUBSCRIPTIONS (Pro + Credits)
            # =========================
            gr.Markdown("## 💎 Account Subscriptions")

            with gr.Accordion(" ", open=False):

                # -------------------------
                # PRO (teacher-facing copy)
                # -------------------------
                with gr.Accordion("💎 Pro (Paid Plan)", open=False):
                    gr.Markdown(
                        "**Pro is for teachers who want complete, classroom-ready materials.**\n\n"
                        "**With Pro you can:**\n"
                        "- ✅ **Export** to **DOCX** and **PowerPoint**\n"
                        "- ✅ **Save** your generated work in **Your Saved Work** (so you can come back later)\n"
                        "- ✅ **Use Templates** (save and reuse your best activities)\n"
                        "- ✅ **View full diagrams and AI-generated images** inside your documents\n\n"
                        "**Free plan:** you can still generate content and preview it in markdown, but **saving, exporting, templates, diagrams, and images are locked**."
                    )

                    upgrade_url = os.getenv("PRO_UPGRADE_URL", "").strip()
                    manage_url  = os.getenv("PRO_MANAGE_URL", "").strip()

                    with gr.Row():
                        activate_pro_btn = gr.Button("🔓 Activate Pro", variant="primary", visible=True)
                        downgrade_pro_btn = gr.Button("🔻 Downgrade to Free", visible=False)
                        pro_status = gr.Textbox(
                            label="Pro status",
                            lines=2,
                            interactive=False
                        )

                    with gr.Row():
                        if upgrade_url:
                            gr.Markdown(f"🌐 **Future checkout link:** [{upgrade_url}]({upgrade_url})")
                        else:
                            gr.Markdown("🌐 **Future checkout link:** Not configured yet.")

                        if manage_url:
                            gr.Markdown(f"⚙️ **Future manage subscription link:** [{manage_url}]({manage_url})")
                        else:
                            gr.Markdown("⚙️ **Future manage subscription link:** Not configured yet.")

                # -------------------------
                # CREDITS
                # -------------------------
                with gr.Accordion("💳 Credits", open=False):
                    gr.Markdown(
                        "Credits are used for certain paid actions (for example, generating images/diagrams).\n"
                        "Your balance is shown below."
                    )

                    credits_btn = gr.Button("Refresh credits")
                    credits_refresh_state_profile = gr.Textbox(label="Balance", lines=2, interactive=False)

                    credits_btn.click(
                        fn=credits_status_text,
                        inputs=[supabase_session],
                        outputs=[credits_refresh_state_profile]
                    )

                    credits_refresh_state_profile.change(
                        fn=lambda v: v,
                        inputs=[credits_refresh_state_profile],
                        outputs=[credits_refresh_state_profile]
                    )

            
            # =========================
            # LEGAL (ACKNOWLEDGEMENTS) — B2 (3 accordions only)
            # =========================
            gr.Markdown("## 📜 Legal")

            with gr.Accordion(" ", open=False):

                gr.Markdown(
                    "Please review and acknowledge these documents. Once acknowledged, it is recorded on your profile."
                )

                legal_up_to_date_note = gr.Markdown(
                    "✅ You are up to date with the current legal requirements."
                )

                # -------------------------
                # Privacy Policy
                # -------------------------
                with gr.Accordion("🛡️ Privacy Policy", open=False):
                    gr.Markdown(
                        "Covers how EduDRAFT STUDIO collects, stores, and protects your information."
                    )

                    privacy_doc_md = gr.Markdown(
                        value=LEGAL_MD_TEXT.get("privacy", ""),
                        elem_classes=["legal_doc_box"]
                    )

                    privacy_ack_status = gr.Markdown("❌ Not acknowledged yet.")
                    privacy_ack_cb = gr.Checkbox(label="I have read and agree to the Privacy Policy", value=False)

                # -------------------------
                # Terms of Use
                # -------------------------
                with gr.Accordion("📘 Terms of Use", open=False):
                    gr.Markdown(
                        "Covers acceptable use of the app, teacher responsibilities, and content review expectations."
                    )

                    terms_doc_md = gr.Markdown(
                        value=LEGAL_MD_TEXT.get("terms", ""),
                        elem_classes=["legal_doc_box"]
                    )

                    terms_ack_status = gr.Markdown("❌ Not acknowledged yet.")
                    terms_ack_cb = gr.Checkbox(label="I have read and agree to the Terms of Use", value=False)

                # -------------------------
                # Terms & Conditions
                # -------------------------
                with gr.Accordion("🧾 Terms & Conditions", open=False):
                    gr.Markdown(
                        "Covers commercial terms for the service (where applicable), including plan rules and credit-related terms."
                    )

                    tcs_doc_md = gr.Markdown(
                        value=LEGAL_MD_TEXT.get("tcs", ""),
                        elem_classes=["legal_doc_box"]
                    )

                    tcs_ack_status = gr.Markdown("❌ Not acknowledged yet.")
                    tcs_ack_cb = gr.Checkbox(label="I have read and agree to the Terms & Conditions", value=False)


                # -------------------------
                # Acknowledgement wiring (user-driven to avoid snapshot loops)
                # -------------------------
                privacy_ack_cb.input(
                    fn=lambda sess, checked: action_ack_legal_doc(sess, "privacy", checked),
                    inputs=[supabase_session, privacy_ack_cb],
                    outputs=[privacy_ack_status, privacy_ack_cb]
                )

                terms_ack_cb.input(
                    fn=lambda sess, checked: action_ack_legal_doc(sess, "terms", checked),
                    inputs=[supabase_session, terms_ack_cb],
                    outputs=[terms_ack_status, terms_ack_cb]
                )

                tcs_ack_cb.input(
                    fn=lambda sess, checked: action_ack_legal_doc(sess, "tcs", checked),
                    inputs=[supabase_session, tcs_ack_cb],
                    outputs=[tcs_ack_status, tcs_ack_cb]
                )

                # -------------------------
                # Debug (developer tool)
                # -------------------------
                with gr.Accordion("🛠 Debug Legal State", open=False):
                    debug_legal_btn = gr.Button("DEBUG LEGAL STATE")
                    debug_legal_out = gr.Textbox(label="Legal Debug Output", lines=15, interactive=False)


            # =========================
            # DELETE / PAUSE ACCOUNT
            # =========================
            gr.Markdown("## 🗑️ Delete Account")

            with gr.Accordion(" ", open=False):

                gr.Markdown(
                    "You’re in control of your account.\n\n"
                    "**Pause Account** temporarily disables your access. You can come back later and reactivate it.\n\n"
                    "**Delete Account** permanently removes your account (and saved data where applicable). This cannot be undone."
                )

                # -------------------------
                # PAUSE
                # -------------------------
                with gr.Accordion("⏸️ Pause / Reactivate", open=False):

                    # --- Pause ---
                    pause_confirm = gr.Checkbox(
                        label="I understand pausing will disable my access until I reactivate.",
                        value=False
                    )
                    pause_btn = gr.Button("⏸️ Pause Account", interactive=False)
                    pause_status = gr.Textbox(label="Pause status", lines=2, interactive=False)

                    gr.Markdown("---")

                    # --- Unpause ---
                    unpause_confirm = gr.Checkbox(
                        label="I confirm I want to reactivate my account.",
                        value=False
                    )
                    unpause_btn = gr.Button("▶️ Reactivate Account", interactive=False)
                    unpause_status = gr.Textbox(label="Reactivate status", lines=2, interactive=False)


                # -------------------------
                # DELETE (PERMANENT)
                # -------------------------
                with gr.Accordion("🗑️ Delete Account (Permanent)", open=False):
                    gr.Markdown(
                        "**Before you delete:**\n"
                        "- This action is **permanent** and cannot be undone.\n"
                        "- You may lose access to saved drafts, templates, and exports.\n"
                        "- If you have an active subscription, you may need to cancel it separately (when billing is enabled).\n"
                        "- Some data may be retained if required by law (for example, payment records)."
                    )

                    del_c1 = gr.Checkbox(label="I understand deleting my account is permanent.", value=False)
                    del_c2 = gr.Checkbox(label="I understand I may lose saved work and templates.", value=False)
                    del_typed = gr.Textbox(label='Type "DELETE" to confirm', placeholder="DELETE", value="")

                    del_btn = gr.Button("🗑️ Delete Account", interactive=False)
                    del_status = gr.Textbox(label="Delete status", lines=2, interactive=False)


            def _profile_template_viewer_lock(sess):
                # 1) Hard lock: paused/delete requested
                locked, reason = _account_lock_status(sess)
                if locked:
                    return (
                        gr.update(interactive=False),
                        gr.update(interactive=False),
                        gr.update(interactive=False),
                        f"{reason} — Template Viewer is disabled."
                    )

                # 2) Plan lock: Free users cannot use Templates
                if not _is_pro_user(sess):
                    return (
                        gr.update(interactive=False),
                        gr.update(interactive=False),
                        gr.update(interactive=False),
                        "🔒 Pro required — Templates are locked on the Free plan."
                    )

                # 3) Allowed
                return (
                    gr.update(interactive=True),
                    gr.update(interactive=True),
                    gr.update(interactive=True),
                    ""
                )

            # -------------------------
            # Button-enable wiring
            # -------------------------
            def _pause_btn_state(confirmed: bool):
                return gr.update(interactive=bool(confirmed))

            def _unpause_btn_state(confirmed: bool):
                return gr.update(interactive=bool(confirmed))

            def _delete_btn_state(c1: bool, c2: bool, typed: str):
                ok = bool(c1) and bool(c2) and ((typed or "").strip().upper() == "DELETE")
                return gr.update(interactive=ok)

            pause_confirm.change(
                fn=_pause_btn_state,
                inputs=[pause_confirm],
                outputs=[pause_btn]
            )

            unpause_confirm.change(
                fn=_unpause_btn_state,
                inputs=[unpause_confirm],
                outputs=[unpause_btn]
            )

            del_c1.change(fn=_delete_btn_state, inputs=[del_c1, del_c2, del_typed], outputs=[del_btn])
            del_c2.change(fn=_delete_btn_state, inputs=[del_c1, del_c2, del_typed], outputs=[del_btn])
            del_typed.change(fn=_delete_btn_state, inputs=[del_c1, del_c2, del_typed], outputs=[del_btn])

            # -------------------------
            # Backend wiring
            # -------------------------
            pause_btn.click(
                fn=action_pause_account,
                inputs=[supabase_session, pause_confirm],
                outputs=[pause_status]
            ).then(
                fn=supa_plan_badge,
                inputs=[supabase_session],
                outputs=[plan_badge]
            ).then(
                fn=_profile_template_viewer_lock,
                inputs=[supabase_session],
                outputs=[profile_templates_refresh_btn, profile_preview_template_btn, profile_load_template_btn, profile_template_status]
            )

            unpause_btn.click(
                fn=action_unpause_account,
                inputs=[supabase_session, unpause_confirm],
                outputs=[unpause_status]
            ).then(
                fn=supa_plan_badge,
                inputs=[supabase_session],
                outputs=[plan_badge]
            ).then(
                fn=_update_export_ui,
                inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            )

            del_btn.click(
                fn=action_delete_account_request,
                inputs=[supabase_session, del_c1, del_c2, del_typed],
                outputs=[del_status]
            ).then(
                fn=_profile_template_viewer_lock,
                inputs=[supabase_session],
                outputs=[profile_templates_refresh_btn, profile_preview_template_btn, profile_load_template_btn, profile_template_status]
            )


            def _saved_work_to_details_html(drafts_count_text: str, drafts_list_text: str) -> str:
                """
                Convert the grouped markdown produced by supa_profile_snapshot() into <details> accordions.
                Input drafts_list_text format is like:

                ### Test / Quiz
                • Title — Subject
                • Title — Subject

                ### Worksheet
                • ...

                Returns HTML string for gr.Markdown to render.
                """
                drafts_count_text = (drafts_count_text or "").strip()
                drafts_list_text = (drafts_list_text or "").strip()

                if not drafts_list_text or drafts_list_text == "(No drafts yet)":
                    # Keep it simple if nothing exists
                    return "(No drafts yet)"

                # Parse into sections
                sections = []
                current_title = None
                current_items = []

                for raw in drafts_list_text.splitlines():
                    line = (raw or "").strip()

                    if line.startswith("### "):
                        # flush previous
                        if current_title is not None:
                            sections.append((current_title, current_items))
                        current_title = line.replace("### ", "", 1).strip()
                        current_items = []
                        continue

                    if line.startswith("• "):
                        current_items.append(line.replace("• ", "", 1).strip())
                        continue

                # flush last
                if current_title is not None:
                    sections.append((current_title, current_items))

                # Build HTML
                out = []
                # We intentionally do NOT show the overall drafts total here,
                # because each section already shows its own count (cleaner UI).

                for title, items in sections:
                    n = len(items or [])
                    out.append(
                        "<details style='margin: 6px 0; padding: 4px 6px; border: 1px solid #ddd; border-radius: 8px;'>"
                    )
                    out.append(f"<summary style='cursor:pointer; font-weight:600;'>{title} ({n})</summary>")

                    if n == 0:
                        out.append("<div style='margin-left: 10px; margin-top: 6px;'>(None yet)</div>")
                    else:
                        out.append("<ul style='margin-top: 6px; margin-left: 18px;'>")
                        for it in items:
                            safe = (it or "").replace("<", "&lt;").replace(">", "&gt;")
                            out.append(f"<li>{safe}</li>")
                        out.append("</ul>")

                    out.append("</details>")

                return "\n".join(out)

            def _templates_to_details_html(templates_count_text: str, templates_list_text: str) -> str:
                """
                Convert grouped markdown produced by supa_profile_snapshot() into <details> accordions.

                It supports nested grouping:
                  ### Subject (n)
                  **▸ Category (k)**
                      • Template
                      • Template

                Returns HTML string for gr.Markdown to render.
                """
                templates_list_text = (templates_list_text or "").strip()

                if not templates_list_text or templates_list_text == "(No templates yet)":
                    return "(No templates yet)"

                import re

                def _strip_trailing_count(label: str) -> str:
                    # Removes a trailing " (123)" if it exists, so we don't show double counts.
                    if not label:
                        return label
                    return re.sub(r"\s*\(\d+\)\s*$", "", label).strip()

                # Parse into subjects -> categories -> items
                subjects = []  # list of (subject_title, categories_dict)
                current_subject = None
                current_cats = {}
                current_cat = None

                def _flush_subject():
                    nonlocal current_subject, current_cats
                    if current_subject is not None:
                        subjects.append((current_subject, current_cats))
                    current_subject = None
                    current_cats = {}

                for raw in templates_list_text.splitlines():
                    line = (raw or "").strip()
                    if not line:
                        continue

                    # Subject header
                    if line.startswith("### "):
                        _flush_subject()
                        current_subject = _strip_trailing_count(line.replace("### ", "", 1).strip())
                        current_cat = None
                        continue

                    # Category header (handles **▸ Cat (n)** or ▸ Cat (n))
                    if line.startswith("**▸ "):
                        cat_title = line.replace("**▸ ", "", 1).strip()
                        if cat_title.endswith("**"):
                            cat_title = cat_title[:-2].strip()
                        current_cat = _strip_trailing_count(cat_title)
                        current_cats.setdefault(current_cat, [])
                        continue

                    if line.startswith("▸ "):
                        current_cat = _strip_trailing_count(line.replace("▸ ", "", 1).strip())
                        current_cats.setdefault(current_cat, [])
                        continue

                    # Template line
                    if line.startswith("• "):
                        item = line.replace("• ", "", 1).strip()
                        if not current_cat:
                            current_cat = "Custom"
                            current_cats.setdefault(current_cat, [])
                        current_cats[current_cat].append(item)
                        continue

                    # If something unexpected appears, treat it as a template item under current category
                    if current_subject is not None:
                        if not current_cat:
                            current_cat = "Custom"
                            current_cats.setdefault(current_cat, [])
                        current_cats[current_cat].append(line)

                _flush_subject()

                # Build HTML
                out = []
                for subj_title, cats in subjects:
                    # total templates in this subject
                    total = sum(len(v or []) for v in (cats or {}).values())

                    out.append(
                        "<details style='margin: 6px 0; padding: 6px 8px; border: 1px solid #ddd; border-radius: 8px;'>"
                    )
                    safe_subj = (subj_title or "").replace("<", "&lt;").replace(">", "&gt;")
                    out.append(f"<summary style='cursor:pointer; font-weight:600;'>{safe_subj} ({total})</summary>")

                    if not cats:
                        out.append("<div style='margin-left: 10px; margin-top: 6px;'>(None yet)</div>")
                    else:
                        # Keep a nice stable order: show categories with content, alphabetically by their display string
                        for cat_title in sorted(cats.keys(), key=lambda s: (s or "").lower()):
                            items = cats.get(cat_title) or []
                            if not items:
                                continue

                            safe_cat = (cat_title or "").replace("<", "&lt;").replace(">", "&gt;")
                            out.append(f"<div style='margin-top: 8px; margin-left: 18px; font-weight:600;'>▸ {safe_cat} ({len(items)})</div>")
                            out.append("<ul style='margin-top: 6px; margin-left: 18px;'>")
                            for it in items:
                                safe = (it or "").replace("<", "&lt;").replace(">", "&gt;")
                                out.append(f"<li>{safe}</li>")
                            out.append("</ul>")

                    out.append("</details>")

                return "\n".join(out)


            # =========================
            # PROFILE: Template Viewer wiring (Preview + Load)
            # =========================

            def _profile_preview_template(sess, template_choice):
                """Preview selected template into read-only boxes (no editor changes)."""
                md, ppt, msg = load_template(sess, template_choice)
                return md or "", ppt or "", msg or ""

            profile_templates_refresh_btn.click(
                fn=refresh_templates_on_open,
                inputs=[supabase_session, profile_template_category_filter],
                outputs=[profile_templates_dd, profile_template_status]
            )

            profile_template_category_filter.change(
                fn=refresh_templates_on_open,
                inputs=[supabase_session, profile_template_category_filter],
                outputs=[profile_templates_dd, profile_template_status]
            )

            profile_preview_template_btn.click(
                fn=_profile_preview_template,
                inputs=[supabase_session, profile_templates_dd],
                outputs=[profile_template_md, profile_template_ppt, profile_template_status]
            )

            profile_load_template_btn.click(
                fn=load_template,
                inputs=[supabase_session, profile_templates_dd],
                outputs=[edited_md, edited_ppt, status]
            ).then(
                fn=_reset_export_guard,
                inputs=[],
                outputs=[export_snooze, export_confirm_state]
            ).then(
                fn=_update_export_ui,
                inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
                outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
            )


            # ---- Wiring ----
            snap_btn.click(
                fn=supa_profile_snapshot,
                inputs=[supabase_session],
                outputs=[
                    # status + banner
                    snap_status,
                    avatar_preview,
                    profile_header,

                    # Account Details (NEW)
                    email_in,
                    first_name_in,
                    last_name_in,
                    country_in,
                    language_in,
                    years_taught_in,
                    school_type_in,
                    curriculum_system_in,
                    subjects_taught_in,
                    assessment_style_in,
                    difficulty_pref_in,
                    spelling_pref_in,
                    marking_style_in,
                    preferred_tone_in,
                    class_size_in,
                    ability_mix_in,
                    confirm_educator_in,
                    confirm_review_in,

                    # Existing
                    display_name_in,
                    notify_weekly,
                    notify_export_done,
                    notify_low_credits,

                    # Saved work
                    drafts_count_out,
                    drafts_list_out,
                    templates_count_out,
                    templates_list_out,

                    # Legal
                    privacy_ack_status,
                    terms_ack_status,
                    tcs_ack_status,
                    privacy_ack_cb,
                    terms_ack_cb,
                    tcs_ack_cb,
                    legal_up_to_date_note,
                ]
            ).then(
                fn=_saved_work_to_details_html,
                inputs=[drafts_count_out, drafts_list_out],
                outputs=[drafts_list_out]
            ).then(
                fn=_templates_to_details_html,
                inputs=[templates_count_out, templates_list_out],
                outputs=[templates_list_out]
            )

            snap_btn.click(
                fn=supa_plan_badge,
                inputs=[supabase_session],
                outputs=[plan_badge]
            ).then(
                fn=_profile_template_viewer_lock,
                inputs=[supabase_session],
                outputs=[profile_templates_refresh_btn, profile_preview_template_btn, profile_load_template_btn, profile_template_status]
            ).then(
                fn=_pro_button_visibility,
                inputs=[supabase_session],
                outputs=[activate_pro_btn, downgrade_pro_btn]
            )

            supabase_session.change(
                fn=supa_profile_snapshot,
                inputs=[supabase_session],
                outputs=[
                    # status + banner
                    snap_status,
                    avatar_preview,
                    profile_header,

                    # Account Details (NEW)
                    email_in,
                    first_name_in,
                    last_name_in,
                    country_in,
                    language_in,
                    years_taught_in,
                    school_type_in,
                    curriculum_system_in,
                    subjects_taught_in,
                    assessment_style_in,
                    difficulty_pref_in,
                    spelling_pref_in,
                    marking_style_in,
                    preferred_tone_in,
                    class_size_in,
                    ability_mix_in,
                    confirm_educator_in,
                    confirm_review_in,

                    # Existing
                    display_name_in,
                    notify_weekly,
                    notify_export_done,
                    notify_low_credits,

                    # Saved work
                    drafts_count_out,
                    drafts_list_out,
                    templates_count_out,
                    templates_list_out,

                    # Legal
                    privacy_ack_status,
                    terms_ack_status,
                    tcs_ack_status,
                    privacy_ack_cb,
                    terms_ack_cb,
                    tcs_ack_cb,
                    legal_up_to_date_note,
                ]
            ).then(
                fn=_saved_work_to_details_html,
                inputs=[drafts_count_out, drafts_list_out],
                outputs=[drafts_list_out]
            ).then(
                fn=_templates_to_details_html,
                inputs=[templates_count_out, templates_list_out],
                outputs=[templates_list_out]
            ).then(
                fn=supa_plan_badge,
                inputs=[supabase_session],
                outputs=[plan_badge]
            ).then(
                fn=_profile_template_viewer_lock,
                inputs=[supabase_session],
                outputs=[profile_templates_refresh_btn, profile_preview_template_btn, profile_load_template_btn, profile_template_status]
            ).then(
                fn=_pro_button_visibility,
                inputs=[supabase_session],
                outputs=[activate_pro_btn, downgrade_pro_btn]
            ).then(
                fn=_apply_templates_plan_lock,
                inputs=[supabase_session],
                outputs=[
                    templates_lock_note,
                    upload_template_confirm,
                    save_analyzed_template_btn,
                    refresh_templates_btn,
                    apply_template_btn,
                    delete_template_btn,
                    ref_attach_btn,
                    ref_load_btn,
                    templates_dropdown,
                    ref_template_dropdown,
                    up_file,
                    load_template_status
                ]
            )

            activate_pro_btn.click(
                fn=action_activate_pro,
                inputs=[supabase_session],
                outputs=[pro_status]
            ).then(
                fn=supa_profile_snapshot,
                inputs=[supabase_session],
                outputs=[
                    # status + banner
                    snap_status,
                    avatar_preview,
                    profile_header,

                    # Account Details (NEW)
                    email_in,
                    first_name_in,
                    last_name_in,
                    country_in,
                    language_in,
                    years_taught_in,
                    school_type_in,
                    curriculum_system_in,
                    subjects_taught_in,
                    assessment_style_in,
                    difficulty_pref_in,
                    spelling_pref_in,
                    marking_style_in,
                    preferred_tone_in,
                    class_size_in,
                    ability_mix_in,
                    confirm_educator_in,
                    confirm_review_in,

                    # Existing
                    display_name_in,
                    notify_weekly,
                    notify_export_done,
                    notify_low_credits,

                    # Saved work
                    drafts_count_out,
                    drafts_list_out,
                    templates_count_out,
                    templates_list_out,

                    # Legal
                    privacy_ack_status,
                    terms_ack_status,
                    tcs_ack_status,
                    privacy_ack_cb,
                    terms_ack_cb,
                    tcs_ack_cb,
                    legal_up_to_date_note,
                ]
            ).then(
                fn=credits_status_text,
                inputs=[supabase_session],
                outputs=[credits_refresh_state_profile]
            ).then(
                fn=_saved_work_to_details_html,
                inputs=[drafts_count_out, drafts_list_out],
                outputs=[drafts_list_out]
            ).then(
                fn=_templates_to_details_html,
                inputs=[templates_count_out, templates_list_out],
                outputs=[templates_list_out]
            ).then(
                fn=supa_plan_badge,
                inputs=[supabase_session],
                outputs=[plan_badge]
            ).then(
                fn=_profile_template_viewer_lock,
                inputs=[supabase_session],
                outputs=[
                    profile_templates_refresh_btn,
                    profile_preview_template_btn,
                    profile_load_template_btn,
                    profile_template_status
                ]
            ).then(
                fn=_apply_drafts_plan_lock,
                inputs=[supabase_session],
                outputs=[
                    drafts_lock_note,
                    draft_search,
                    refresh_btn,
                    drafts_dd,
                    versions_dd,
                    load_btn,
                    rename_new_title,
                    rename_btn,
                    delete_version_btn,
                    delete_draft_btn,
                    library_status,
                ]
            ).then(
                fn=_apply_templates_plan_lock,
                inputs=[supabase_session],
                outputs=[
                    templates_lock_note,
                    upload_template_confirm,
                    save_analyzed_template_btn,
                    refresh_templates_btn,
                    apply_template_btn,
                    delete_template_btn,
                    ref_attach_btn,
                    ref_load_btn,
                    templates_dropdown,
                    ref_template_dropdown,
                    up_file,
                    load_template_status,
                ]
            ).then(
                fn=_pro_button_visibility,
                inputs=[supabase_session],
                outputs=[activate_pro_btn, downgrade_pro_btn]
            )

            downgrade_pro_btn.click(
                fn=action_deactivate_pro,
                inputs=[supabase_session],
                outputs=[pro_status]
            ).then(
                fn=supa_profile_snapshot,
                inputs=[supabase_session],
                outputs=[
                    # status + banner
                    snap_status,
                    avatar_preview,
                    profile_header,

                    # Account Details (NEW)
                    email_in,
                    first_name_in,
                    last_name_in,
                    country_in,
                    language_in,
                    years_taught_in,
                    school_type_in,
                    curriculum_system_in,
                    subjects_taught_in,
                    assessment_style_in,
                    difficulty_pref_in,
                    spelling_pref_in,
                    marking_style_in,
                    preferred_tone_in,
                    class_size_in,
                    ability_mix_in,
                    confirm_educator_in,
                    confirm_review_in,

                    # Existing
                    display_name_in,
                    notify_weekly,
                    notify_export_done,
                    notify_low_credits,

                    # Saved work
                    drafts_count_out,
                    drafts_list_out,
                    templates_count_out,
                    templates_list_out,

                    # Legal
                    privacy_ack_status,
                    terms_ack_status,
                    tcs_ack_status,
                    privacy_ack_cb,
                    terms_ack_cb,
                    tcs_ack_cb,
                    legal_up_to_date_note,
                ]
            ).then(
                fn=credits_status_text,
                inputs=[supabase_session],
                outputs=[credits_refresh_state_profile]
            ).then(
                fn=_saved_work_to_details_html,
                inputs=[drafts_count_out, drafts_list_out],
                outputs=[drafts_list_out]
            ).then(
                fn=_templates_to_details_html,
                inputs=[templates_count_out, templates_list_out],
                outputs=[templates_list_out]
            ).then(
                fn=supa_plan_badge,
                inputs=[supabase_session],
                outputs=[plan_badge]
            ).then(
                fn=_profile_template_viewer_lock,
                inputs=[supabase_session],
                outputs=[
                    profile_templates_refresh_btn,
                    profile_preview_template_btn,
                    profile_load_template_btn,
                    profile_template_status
                ]
            ).then(
                fn=_apply_drafts_plan_lock,
                inputs=[supabase_session],
                outputs=[
                    drafts_lock_note,
                    draft_search,
                    refresh_btn,
                    drafts_dd,
                    versions_dd,
                    load_btn,
                    rename_new_title,
                    rename_btn,
                    delete_version_btn,
                    delete_draft_btn,
                    library_status,
                ]
            ).then(
                fn=_apply_templates_plan_lock,
                inputs=[supabase_session],
                outputs=[
                    templates_lock_note,
                    upload_template_confirm,
                    save_analyzed_template_btn,
                    refresh_templates_btn,
                    apply_template_btn,
                    delete_template_btn,
                    ref_attach_btn,
                    ref_load_btn,
                    templates_dropdown,
                    ref_template_dropdown,
                    up_file,
                    load_template_status,
                ]
            ).then(
                fn=_pro_button_visibility,
                inputs=[supabase_session],
                outputs=[activate_pro_btn, downgrade_pro_btn]
            )

            if False:
                login_btn.click(
                    fn=auth_login_with_legal,
                    inputs=[auth_email, auth_password, legal_privacy_cb, legal_terms_cb, legal_tcs_cb],
                    outputs=[auth_status, supabase_session]
                ).then(
                    fn=_tab_after_auth,
                    inputs=[auth_status, supabase_session],
                    outputs=[current_tab, is_logged_in]
                ).then(
                    fn=lambda tab: gr.update(selected=tab),
                    inputs=[current_tab],
                    outputs=[tabs]
                ).then(
                    fn=supa_global_banner_payload,
                    inputs=[supabase_session],
                    outputs=[global_banner, global_avatar, global_name, global_signed]
                ).then(
                    fn=supa_plan_badge,
                    inputs=[supabase_session],
                    outputs=[plan_badge]
                ).then(
                    fn=supa_profile_snapshot,
                    inputs=[supabase_session],
                    outputs=[
                        # status + banner
                        snap_status,
                        avatar_preview,
                        profile_header,

                        # Account Details (NEW)
                        email_in,
                        first_name_in,
                        last_name_in,
                        country_in,
                        language_in,
                        years_taught_in,
                        school_type_in,
                        curriculum_system_in,
                        subjects_taught_in,
                        assessment_style_in,
                        difficulty_pref_in,
                        spelling_pref_in,
                        marking_style_in,
                        preferred_tone_in,
                        class_size_in,
                        ability_mix_in,
                        confirm_educator_in,
                        confirm_review_in,

                        # Existing
                        display_name_in,
                        notify_weekly,
                        notify_export_done,
                        notify_low_credits,

                        # Saved work
                        drafts_count_out,
                        drafts_list_out,
                        templates_count_out,
                        templates_list_out,

                        # Legal
                        privacy_ack_status,
                        terms_ack_status,
                        tcs_ack_status,
                        privacy_ack_cb,
                        terms_ack_cb,
                        tcs_ack_cb,
                        legal_up_to_date_note,
                    ]
                ).then(
                    fn=credits_status_text,
                    inputs=[supabase_session],
                    outputs=[credits_refresh_state_profile]
                ).then(
                    fn=_saved_work_to_details_html,
                    inputs=[drafts_count_out, drafts_list_out],
                    outputs=[drafts_list_out]
                ).then(
                    fn=_templates_to_details_html,
                    inputs=[templates_count_out, templates_list_out],
                    outputs=[templates_list_out]
                ).then(
                    fn=_profile_template_viewer_lock,
                    inputs=[supabase_session],
                    outputs=[profile_templates_refresh_btn, profile_preview_template_btn, profile_load_template_btn, profile_template_status]
                )

            avatar_save_btn.click(
                fn=supa_set_avatar,
                inputs=[supabase_session, avatar_upload],
                outputs=[avatar_status, avatar_preview]
            ).then(
                fn=supa_profile_banner,
                inputs=[supabase_session],
                outputs=[banner_avatar, banner_name, banner_signed]
            ).then(
                fn=supa_global_banner_payload,
                inputs=[supabase_session],
                outputs=[global_banner, global_avatar, global_name, global_signed]
            )

            avatar_reload_btn.click(
                fn=supa_load_avatar,
                inputs=[supabase_session],
                outputs=[avatar_status, avatar_preview]
            ).then(
                fn=supa_profile_banner,
                inputs=[supabase_session],
                outputs=[banner_avatar, banner_name, banner_signed]
            ).then(
                fn=supa_global_banner_payload,
                inputs=[supabase_session],
                outputs=[global_banner, global_avatar, global_name, global_signed]
            )

            save_profile_btn.click(
                fn=supa_save_profile_v2,
                inputs=[
                    supabase_session,

                    # New Account Details fields
                    first_name_in,
                    last_name_in,
                    email_in,
                    country_in,
                    language_in,
                    years_taught_in,
                    school_type_in,
                    curriculum_system_in,
                    subjects_taught_in,
                    assessment_style_in,
                    difficulty_pref_in,
                    spelling_pref_in,
                    marking_style_in,
                    preferred_tone_in,
                    class_size_in,
                    ability_mix_in,
                    confirm_educator_in,
                    confirm_review_in,

                    # Existing fields (still used)
                    display_name_in,
                    notify_weekly,
                    notify_export_done,
                    notify_low_credits
                ],
                outputs=[save_profile_status]
            ).then(
                fn=supa_profile_summary_md,
                inputs=[supabase_session],
                outputs=[saved_profile_summary]
            ).then(
                fn=supa_profile_banner,
                inputs=[supabase_session],
                outputs=[banner_avatar, banner_name, banner_signed]
            ).then(
                fn=supa_global_banner_payload,
                inputs=[supabase_session],
                outputs=[global_banner, global_avatar, global_name, global_signed]
            )

            profile_logout_btn.click(
                fn=auth_logout,
                inputs=[supabase_session],
                outputs=[save_profile_status, supabase_session]
            ).then(
                fn=lambda: ("Login", False),
                inputs=[],
                outputs=[current_tab, is_logged_in]
            ).then(
                fn=lambda tab: gr.update(selected=tab),
                inputs=[current_tab],
                outputs=[tabs]
            ).then(
                fn=None,
                inputs=[],
                outputs=[browser_session_bridge],
                queue=False,
                show_progress="hidden",
                js=f"""
                () => {{
                    const key = "{BROWSER_SESSION_STORAGE_KEY}";
                    try {{
                        window.localStorage.removeItem(key);
                    }} catch (e) {{}}
                    return "";
                }}
                """
            )

            debug_legal_btn.click(
                fn=debug_legal_state,
                inputs=[supabase_session],
                outputs=[debug_legal_out]
            )

        def _profile_banner_on_login(logged_in, sess):
            if not logged_in:
                return (None, "**(not loaded)**", "_Not signed in._")
            return supa_profile_banner(sess)

        is_logged_in.change(
            fn=_profile_banner_on_login,
            inputs=[is_logged_in, supabase_session],
            outputs=[banner_avatar, banner_name, banner_signed]
        )


        # =========================
        # PROFILE EMAIL: set email field ONCE after login
        # =========================
        def _email_on_login(logged_in, sess):
            if not logged_in:
                return ""
            return _email_from_session(sess) or ""

        is_logged_in.change(
            fn=_email_on_login,
            inputs=[is_logged_in, supabase_session],
            outputs=[email_in]
        )

        supabase_session.change(
            fn=_profile_template_viewer_lock,
            inputs=[supabase_session],
            outputs=[profile_templates_refresh_btn, profile_preview_template_btn, profile_load_template_btn, profile_template_status]
        )


        # =========================
        # Generate button wiring (HARD LOCK aware)
        # =========================
        input_mode.change(
            fn=_compute_generate_ui_state,
            inputs=[input_mode, audio_in, typed_prompt, live_transcript, supabase_session, course, other_subject, education_level],
            outputs=[gen_btn]
        )
        audio_in.change(
            fn=_compute_generate_ui_state,
            inputs=[input_mode, audio_in, typed_prompt, live_transcript, supabase_session, course, other_subject, education_level],
            outputs=[gen_btn]
        )
        typed_prompt.change(
            fn=_compute_generate_ui_state,
            inputs=[input_mode, audio_in, typed_prompt, live_transcript, supabase_session, course, other_subject, education_level],
            outputs=[gen_btn]
        )
        course.change(
            fn=_compute_generate_ui_state,
            inputs=[input_mode, audio_in, typed_prompt, live_transcript, supabase_session, course, other_subject, education_level],
            outputs=[gen_btn]
        )
        other_subject.change(
            fn=_compute_generate_ui_state,
            inputs=[input_mode, audio_in, typed_prompt, live_transcript, supabase_session, course, other_subject, education_level],
            outputs=[gen_btn]
        )
        supabase_session.change(
            fn=_compute_generate_ui_state,
            inputs=[input_mode, audio_in, typed_prompt, live_transcript, supabase_session, course, other_subject, education_level],
            outputs=[gen_btn]
        )


        # =========================
        # Export wiring (FIXED)
        # =========================

        export_btn.click(
            fn=action_export_files,
            inputs=[supabase_session, export_confirm_cb, edited_md, edited_ppt],
            outputs=[docx_file, pptx_file, status]
        ).then(
            fn=_after_export_snooze,
            inputs=[status],
            outputs=[export_snooze, export_confirm_state]
        ).then(
            fn=_update_export_ui,
            inputs=[edited_md, edited_ppt, export_snooze, export_confirm_state, supabase_session],
            outputs=[export_cost_note, export_confirm_cb, export_btn, save_version_btn]
        ).then(
            fn=credits_status_text,
            inputs=[supabase_session],
            outputs=[credits_refresh_state_profile]   # keep this
        )


        # =========================
        # ADMIN TAB (diagnostics)
        # =========================
        with gr.TabItem("Admin", id="Admin", visible=False) as admin_tab:
            gr.Markdown("## 🛠️ Admin / Diagnostics")

            gr.Markdown("### 🔐 Supabase secrets check")
            gr.Textbox(value=SUPABASE_DEBUG, label="What the app sees (masked)", interactive=False)

            gr.Markdown("### 🔌 Supabase health check")
            test_btn = gr.Button("Test Supabase connection")
            test_out = gr.Textbox(label="Supabase status")
            test_btn.click(fn=supabase_healthcheck, inputs=[], outputs=test_out)

            gr.Markdown("### 🧪 Image-Gen Self-Check")
            imgcheck_btn = gr.Button("Run image-gen self-check")
            imgcheck_status = gr.Textbox(label="Self-check result", lines=10, interactive=False)
            imgcheck_preview = gr.Image(label="Preview (if successful)")

            imgcheck_btn.click(
                fn=action_image_gen_selfcheck,
                inputs=[],
                outputs=[imgcheck_status, imgcheck_preview]
            )

            gr.Markdown("### 👤 Default Avatar Control (Admin)")
            gr.Markdown(
                "Upload or replace the system default avatar used for users who have not uploaded their own profile image yet."
            )

            default_avatar_file = gr.File(
                label="Upload default avatar",
                file_types=[".png", ".jpg", ".jpeg", ".webp"]
            )

            with gr.Row():
                default_avatar_save_btn = gr.Button("💾 Save default avatar")
                default_avatar_refresh_btn = gr.Button("🔄 Refresh preview")

            default_avatar_status = gr.Textbox(
                label="Default avatar status",
                lines=2,
                interactive=False
            )

            default_avatar_preview = gr.Image(
                label="Default avatar preview",
                interactive=False
            )

            browser_session_debug = gr.Textbox(
                value="",
                visible=True,
                label="Browser Session Debug",
                lines=6,
                interactive=False
            )
            read_browser_session_btn = gr.Button("Read Browser Session Debug")

            gr.Markdown("### 🧬 BS5 / template_engine Live Debug")

            bs5_debug_json = gr.Textbox(
                value="",
                label="Latest BS5 debug snapshot",
                lines=30,
                interactive=False
            )

            bs5_debug_refresh_btn = gr.Button("Refresh BS5 debug from current editor/template")

            def _admin_load_default_avatar_preview(sess):
                if not sess:
                    return "❌ Not signed in.", None
                if not _is_admin_user(sess):
                    return "❌ Admin access required.", None

                signed_url = _signed_storage_url(
                    DEFAULT_AVATAR_BUCKET,
                    DEFAULT_AVATAR_PATH,
                    expires_in=300
                )

                if not signed_url:
                    return "ℹ️ No default avatar uploaded yet.", None

                return "✅ Default avatar loaded.", signed_url

            def _admin_capture_bs5_debug(current_draft_md, template_bundle, year_level_value, subject_value):
                try:
                    from template_engine import debug_template_engine_snapshot

                    clean_year_level = (year_level_value or "").strip()
                    clean_subject = (subject_value or "").strip()

                    # BS5 RULE:
                    # Do not let the default UI year leak back into debug snapshots.
                    if clean_year_level == "Year 11":
                        clean_year_level = ""

                    snapshot = debug_template_engine_snapshot(
                        current_draft_md=current_draft_md,
                        template_bundle=template_bundle,
                        year_level=clean_year_level,
                        subject=clean_subject
                    )
                    return snapshot, snapshot
                except Exception as e:
                    msg = (
                        '{\n'
                        f'  "status": "error",\n'
                        f'  "error_type": "{type(e).__name__}",\n'
                        f'  "error_message": "{str(e)}"\n'
                        '}'
                    )
                    return msg, msg

            default_avatar_save_btn.click(
                fn=supa_set_default_avatar,
                inputs=[supabase_session, default_avatar_file],
                outputs=[default_avatar_status, default_avatar_preview]
            )

            default_avatar_refresh_btn.click(
                fn=_admin_load_default_avatar_preview,
                inputs=[supabase_session],
                outputs=[default_avatar_status, default_avatar_preview]
            )

            bs5_debug_refresh_btn.click(
                fn=_admin_capture_bs5_debug,
                inputs=[edited_md, template_bundle_state, shared_year_level, shared_subject],
                outputs=[bs5_debug_json, bs5_debug_state]
            )

            bs5_debug_state.change(
                fn=lambda x: x or "ℹ️ No BS5 debug snapshot captured yet.",
                inputs=[bs5_debug_state],
                outputs=[bs5_debug_json]
            )

            gr.Markdown("### 📜 Legal Docs Control (Admin)")

            legal_view_md = gr.Markdown("Click refresh to load versions from DB.")
            legal_refresh_btn = gr.Button("🔄 Refresh legal_config")

            with gr.Row():
                admin_privacy_ver = gr.Textbox(label="privacy current_version", value="")
                admin_terms_ver   = gr.Textbox(label="terms current_version", value="")
                admin_tcs_ver     = gr.Textbox(label="tcs current_version", value="")

            legal_save_btn = gr.Button("💾 Save versions (admin)")
            legal_save_status = gr.Textbox(label="Save status", lines=2, interactive=False)

            def _admin_prefill_versions(sess):
                if not sess:
                    return "❌ Not signed in.", "", "", ""
                try:
                    sb = _sb_authed_from_session(sess)
                    req = _get_required_legal_versions(sb)
                    md = _admin_legal_fetch(sess)
                    return md, req.get("privacy",""), req.get("terms",""), req.get("tcs","")
                except Exception as e:
                    return f"❌ Prefill failed: {type(e).__name__}: {e}", "", "", ""

            legal_refresh_btn.click(
                fn=_admin_prefill_versions,
                inputs=[supabase_session],
                outputs=[legal_view_md, admin_privacy_ver, admin_terms_ver, admin_tcs_ver]
            )

            legal_save_btn.click(
                fn=_admin_legal_update,
                inputs=[supabase_session, admin_privacy_ver, admin_terms_ver, admin_tcs_ver],
                outputs=[legal_save_status]
            ).then(
                fn=_admin_prefill_versions,
                inputs=[supabase_session],
                outputs=[legal_view_md, admin_privacy_ver, admin_terms_ver, admin_tcs_ver]
            )

            gr.Markdown("## 📊 Learning & Analytics Dashboard")
            with gr.Row():
                refresh_analytics_btn = gr.Button("Refresh Analytics")
            analytics_output = gr.Markdown("Click refresh to load analytics.")
            
            def _get_analytics_report(sess):
                if not sess:
                    return "Not logged in."
                try:
                    sb = _sb_authed_from_session(sess)
                    corr = sb.table("template_corrections").select("decision_name,teacher_decision,document_type").execute()
                    corr_data = corr.data or []
                    total = len(corr_data)
                    if total == 0:
                        return "No teacher corrections recorded yet."
                    from collections import Counter
                    overrides = Counter()
                    for c in corr_data:
                        if c.get("teacher_decision") is not None:
                            overrides[c.get("decision_name")] += 1
                    top = overrides.most_common(5)
                    top_features = "\n".join([f"- **{f}**: {c} overrides" for f, c in top])
                    doc_counts = Counter([c.get("document_type") or "unknown" for c in corr_data])
                    doc_lines = "\n".join([f"- **{dt}**: {count}" for dt, count in doc_counts.items()])
                    rules = sb.table("learning_rules").select("*").execute()
                    rules_data = rules.data or []
                    rules_lines = []
                    for r in rules_data[:10]:
                        total_dec = r.get("total_decisions", 0)
                        over_cnt = r.get("override_count", 0)
                        suggested = "apply" if r.get("suggested_apply") else "suppress"
                        rules_lines.append(f"- **{r['rule_key']}**: decisions={total_dec}, overrides={over_cnt}, suggests={suggested}")
                    rules_text = "\n".join(rules_lines) if rules_lines else "No learning rules yet."
                    return f"""### Analytics Summary
**Total teacher corrections recorded:** {total}
**Most overridden features:**
{top_features}
**Corrections by document type:**
{doc_lines}
**Active learning rules (aggregated):**
{rules_text}
*Note: Rules are applied after 5 decisions and when override rate > 80%*
"""
                except Exception as e:
                    return f"Error: {e}"
            
            refresh_analytics_btn.click(
                fn=_get_analytics_report,
                inputs=[supabase_session],
                outputs=[analytics_output]
            )

        supabase_session.change(
            fn=_public_vs_app_tab_visibility,
            inputs=[supabase_session],
            outputs=[
                login_tab,
                home_tab,
                signup_tab,
                login2_tab,
                notifications_tab,
                workspace_tab,
                drafts_tab,
                templates_tab,
                profile_tab,
            ]
        )

        supabase_session.change(
            fn=_notifications_legal_ui_snapshot,
            inputs=[supabase_session],
            outputs=[
                notifications_privacy_ack_status,
                notifications_terms_ack_status,
                notifications_tcs_ack_status,
                notifications_privacy_acc,
                notifications_terms_acc,
                notifications_tcs_acc,
            ]
        )

        notifications_privacy_ack_cb.input(
            fn=lambda sess, checked: _notifications_ack_refresh(sess, "privacy", checked),
            inputs=[supabase_session, notifications_privacy_ack_cb],
            outputs=[
                notifications_privacy_ack_status,
                notifications_terms_ack_status,
                notifications_tcs_ack_status,
                notifications_privacy_acc,
                notifications_terms_acc,
                notifications_tcs_acc,
                notifications_privacy_ack_cb,
                current_tab,
                is_logged_in,
            ]
        ).then(
            fn=lambda tab: gr.update(selected=tab),
            inputs=[current_tab],
            outputs=[tabs]
        ).then(
            fn=_public_vs_app_tab_visibility,
            inputs=[supabase_session],
            outputs=[
                login_tab,
                home_tab,
                signup_tab,
                login2_tab,
                notifications_tab,
                workspace_tab,
                drafts_tab,
                templates_tab,
                profile_tab,
            ]
        ).then(
            fn=_profile_legal_status_snapshot,
            inputs=[supabase_session],
            outputs=[
                privacy_ack_status,
                terms_ack_status,
                tcs_ack_status,
            ]
        )

        notifications_terms_ack_cb.input(
            fn=lambda sess, checked: _notifications_ack_refresh(sess, "terms", checked),
            inputs=[supabase_session, notifications_terms_ack_cb],
            outputs=[
                notifications_privacy_ack_status,
                notifications_terms_ack_status,
                notifications_tcs_ack_status,
                notifications_privacy_acc,
                notifications_terms_acc,
                notifications_tcs_acc,
                notifications_terms_ack_cb,
                current_tab,
                is_logged_in,
            ]
        ).then(
            fn=lambda tab: gr.update(selected=tab),
            inputs=[current_tab],
            outputs=[tabs]
        ).then(
            fn=_public_vs_app_tab_visibility,
            inputs=[supabase_session],
            outputs=[
                login_tab,
                home_tab,
                signup_tab,
                login2_tab,
                notifications_tab,
                workspace_tab,
                drafts_tab,
                templates_tab,
                profile_tab,
            ]
        ).then(
            fn=_profile_legal_status_snapshot,
            inputs=[supabase_session],
            outputs=[
                privacy_ack_status,
                terms_ack_status,
                tcs_ack_status,
            ]
        )

        notifications_tcs_ack_cb.input(
            fn=lambda sess, checked: _notifications_ack_refresh(sess, "tcs", checked),
            inputs=[supabase_session, notifications_tcs_ack_cb],
            outputs=[
                notifications_privacy_ack_status,
                notifications_terms_ack_status,
                notifications_tcs_ack_status,
                notifications_privacy_acc,
                notifications_terms_acc,
                notifications_tcs_acc,
                notifications_tcs_ack_cb,
                current_tab,
                is_logged_in,
            ]
        ).then(
            fn=lambda tab: gr.update(selected=tab),
            inputs=[current_tab],
            outputs=[tabs]
        ).then(
            fn=_public_vs_app_tab_visibility,
            inputs=[supabase_session],
            outputs=[
                login_tab,
                home_tab,
                signup_tab,
                login2_tab,
                notifications_tab,
                workspace_tab,
                drafts_tab,
                templates_tab,
                profile_tab,
            ]
        ).then(
            fn=_profile_legal_status_snapshot,
            inputs=[supabase_session],
            outputs=[
                privacy_ack_status,
                terms_ack_status,
                tcs_ack_status,
            ]
        )
        
        supabase_session.change(
            fn=_admin_tab_visibility,
            inputs=[supabase_session],
            outputs=[admin_tab]
        )

        supabase_session.change(
            fn=None,
            inputs=[supabase_session],
            outputs=[browser_session_bridge],
            queue=False,
            show_progress="hidden",
            js=f"""
            (sess) => {{
                const key = "{BROWSER_SESSION_STORAGE_KEY}";
                try {{
                    if (
                        sess &&
                        typeof sess === "object" &&
                        sess.access_token &&
                        sess.refresh_token &&
                        sess.user_id
                    ) {{
                        const blob = JSON.stringify(sess);
                        window.localStorage.setItem(key, blob);
                        return blob;
                    }}

                    window.localStorage.removeItem(key);
                    return "";
                }} catch (e) {{
                    try {{
                        window.localStorage.removeItem(key);
                    }} catch (_) {{}}
                    return "";
                }}
            }}
            """
        )

        # =========================
        # ADMIN TAB VISIBILITY WIRING
        # =========================

        # After sign-up attempt, refresh Admin visibility from the real session
        signup_create_btn.click(
            fn=_admin_tab_visibility,
            inputs=[supabase_session],
            outputs=[admin_tab]
        )

        # After sign-up logout, hide Admin again
        signup_logout_btn.click(
            fn=_admin_tab_visibility,
            inputs=[supabase_session],
            outputs=[admin_tab]
        )

        # After login, refresh Admin visibility from the real session
        login_go_btn.click(
            fn=_admin_tab_visibility,
            inputs=[supabase_session],
            outputs=[admin_tab]
        )

        # After login-tab logout, hide Admin again
        login_logout_btn.click(
            fn=_admin_tab_visibility,
            inputs=[supabase_session],
            outputs=[admin_tab]
        )

    demo.load(
        fn=_restore_browser_session_with_debug,
        inputs=[browser_session_bridge],
        outputs=[supabase_session, current_tab, is_logged_in, browser_session_debug],
        queue=False,
        show_progress="hidden",
        js=f"""
        () => {{
            const key = "{BROWSER_SESSION_STORAGE_KEY}";
            try {{
                return [window.localStorage.getItem(key) || ""];
            }} catch (e) {{
                return [""];
            }}
        }}
        """
    ).then(
        fn=_public_vs_app_tab_visibility,
        inputs=[supabase_session],
        outputs=[
            login_tab,
            home_tab,
            signup_tab,
            login2_tab,
            notifications_tab,
            workspace_tab,
            drafts_tab,
            templates_tab,
            profile_tab,
        ],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=lambda tab: gr.update(selected=tab),
        inputs=[current_tab],
        outputs=[tabs],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=supa_profile_snapshot,
        inputs=[supabase_session],
        outputs=[
            snap_status,
            avatar_preview,
            profile_header,

            email_in,
            first_name_in,
            last_name_in,
            country_in,
            language_in,
            years_taught_in,
            school_type_in,
            curriculum_system_in,
            subjects_taught_in,
            assessment_style_in,
            difficulty_pref_in,
            spelling_pref_in,
            marking_style_in,
            preferred_tone_in,
            class_size_in,
            ability_mix_in,
            confirm_educator_in,
            confirm_review_in,

            display_name_in,
            notify_weekly,
            notify_export_done,
            notify_low_credits,

            drafts_count_out,
            drafts_list_out,
            templates_count_out,
            templates_list_out,

            privacy_ack_status,
            terms_ack_status,
            tcs_ack_status,
            privacy_ack_cb,
            terms_ack_cb,
            tcs_ack_cb,
            legal_up_to_date_note,
        ],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=_saved_work_to_details_html,
        inputs=[drafts_count_out, drafts_list_out],
        outputs=[drafts_list_out],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=_templates_to_details_html,
        inputs=[templates_count_out, templates_list_out],
        outputs=[templates_list_out],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=supa_plan_badge,
        inputs=[supabase_session],
        outputs=[plan_badge],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=_profile_template_viewer_lock,
        inputs=[supabase_session],
        outputs=[profile_templates_refresh_btn, profile_preview_template_btn, profile_load_template_btn, profile_template_status],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=_pro_button_visibility,
        inputs=[supabase_session],
        outputs=[activate_pro_btn, downgrade_pro_btn],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=_apply_templates_plan_lock,
        inputs=[supabase_session],
        outputs=[
            templates_lock_note,
            upload_template_confirm,
            save_analyzed_template_btn,
            refresh_templates_btn,
            apply_template_btn,
            delete_template_btn,
            ref_attach_btn,
            ref_load_btn,
            templates_dropdown,
            ref_template_dropdown,
            up_file,
            load_template_status
        ],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=supa_profile_banner,
        inputs=[supabase_session],
        outputs=[banner_avatar, banner_name, banner_signed],
        queue=False,
        show_progress="hidden"
    ).then(
        fn=supa_global_banner_payload,
        inputs=[supabase_session],
        outputs=[global_banner, global_avatar, global_name, global_signed],
        queue=False,
        show_progress="hidden"
    )

    # =========================
    # SAVE wiring (Workspace)
    # =========================
    save_version_btn.click(
    fn=supa_save_from_editor,
    inputs=[
        supabase_session,
        draft_name,
        draft_subject_state,
        course_stream,
        education_level,
        country,
        state_province,
        year_level,
        course,
        output_type,      # ✅ NEW (must be here)
        edited_md,
        edited_ppt,
        current_draft_id
    ],
        outputs=[status, current_draft_id, current_version]
    ).then(
        fn=get_rate_limit_display,
        inputs=[supabase_session],
        outputs=[rate_limit_display]
    )


import socket

preferred_port = int(os.getenv("GRADIO_SERVER_PORT", "7861"))
fallback_ports = [preferred_port, 7862, 7863, 7864, 7865]

def _port_is_free(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        return s.connect_ex(("127.0.0.1", port)) != 0

launch_port = None
for port in fallback_ports:
    if _port_is_free(port):
        launch_port = port
        break

if launch_port is None:
    raise OSError(f"No free Gradio port found in: {fallback_ports}")

print(f"Launching Gradio on port {launch_port}")

demo.launch(
    server_name="0.0.0.0",
    server_port=launch_port,
    share=True,
    prevent_thread_lock=False
)



