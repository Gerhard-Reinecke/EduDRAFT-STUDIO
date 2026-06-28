# ======================================================================================
# rate_limit.py
# ======================================================================================
# Module: Usage Tracking & Rate Limiting Control Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module manages per-user request limits to control system usage, prevent abuse,
# and protect operational costs associated with LLM and transcription services.
#
# It enforces daily limits on key actions (e.g., generation and transcription),
# tracks usage in Supabase, and provides real-time feedback to the user interface.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Rate Limit Enforcement
#    - Validates whether a user can perform an action (e.g., "generate", "transcribe")
#    - Applies daily request caps per user
#    - Blocks requests when limits are exceeded
#
# 2. Usage Tracking (Per User, Per Day)
#    - Stores request counts in Supabase (user_usage table)
#    - Tracks usage by user_id and UTC date
#    - Automatically initializes usage records on first request of the day
#
# 3. Dynamic Limit Handling
#    - Supports different limits for different action types:
#        • Generation requests (LLM usage)
#        • Transcription requests (audio processing)
#    - Allows easy adjustment of limits via constants
#
# 4. User Feedback Integration
#    - Returns clear, real-time status messages (e.g., "Request 5/50 today")
#    - Provides UI-friendly usage summaries
#    - Displays remaining quota and limit status
#
# 5. Fault Tolerance
#    - Fails open (allows request) if rate limit check encounters an error
#    - Logs errors without interrupting user workflow
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - Supabase → persistent usage tracking (user_usage table)
# - auth.py → session validation and user identification
# - datetime → UTC-based daily tracking
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - User-based isolation (each user tracked independently)
# - UTC-normalised tracking (consistent across regions)
# - Fail-safe behaviour (never block due to system error)
# - Lightweight and fast (minimal overhead per request)
# - Transparent feedback (clear communication to users)
#
# --------------------------------------------------------------------------------------
# LIMIT STRUCTURE
# --------------------------------------------------------------------------------------
# - DAILY_LIMIT_GENERATE:
#     Maximum number of document generation requests per day
#
# - DAILY_LIMIT_TRANSCRIBE:
#     Maximum number of audio transcription requests per day
#
# These values define the operational cost boundaries of the system.
#
# --------------------------------------------------------------------------------------
# DATABASE CONTRACT
# --------------------------------------------------------------------------------------
# Table: user_usage
# Fields:
#   - id (primary key)
#   - user_id (linked to authenticated user)
#   - date (UTC date, ISO format)
#   - request_count (integer)
#
# One record per user per day.
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This module is critical for cost control and system scalability
# - Any changes here directly impact monetisation and pricing models
# - Current implementation uses a simple request counter (not token-based)
# - Can be extended in future to support:
#     • tier-based limits (Free vs Pro)
#     • token-based billing
#     • rolling time windows instead of daily resets
#
# ======================================================================================


from datetime import datetime, timezone

from auth import _require_session
from config import supabase

DAILY_LIMIT_GENERATE = 50
DAILY_LIMIT_TRANSCRIBE = 100


def check_rate_limit(session_state, action: str = "generate") -> tuple[bool, str]:
    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return False, f"❌ {err}"

    try:
        today = datetime.now(timezone.utc).date().isoformat()

        res = supabase.table("user_usage") \
            .select("*") \
            .eq("user_id", user_id) \
            .eq("date", today) \
            .limit(1) \
            .execute()

        rows = getattr(res, "data", None) or []

        if not rows:
            supabase.table("user_usage").insert({
                "user_id": user_id,
                "date": today,
                "request_count": 1
            }).execute()
            limit = DAILY_LIMIT_GENERATE if action == "generate" else DAILY_LIMIT_TRANSCRIBE
            return True, f"✅ First request today. Rate limit: 1/{limit}"

        current_count = rows[0].get("request_count", 0) + 1
        limit = DAILY_LIMIT_GENERATE if action == "generate" else DAILY_LIMIT_TRANSCRIBE

        if current_count > limit:
            return False, f"❌ Daily limit reached ({limit} requests). Please try again tomorrow."

        supabase.table("user_usage") \
            .update({"request_count": current_count}) \
            .eq("id", rows[0]["id"]) \
            .execute()

        return True, f"✅ Request {current_count}/{limit} today"

    except Exception as e:
        print(f"Rate limit error: {e}")
        return True, "⚠️ Rate limit check failed, proceeding anyway."


def get_rate_limit_display(session_state):
    if not session_state:
        return "Not logged in. Rate limits apply after login."

    access_token, refresh_token, user_id, err = _require_session(session_state)
    if err:
        return f"❌ {err}"

    try:
        today = datetime.now(timezone.utc).date().isoformat()
        res = supabase.table("user_usage") \
            .select("request_count") \
            .eq("user_id", user_id) \
            .eq("date", today) \
            .limit(1) \
            .execute()

        rows = getattr(res, "data", None) or []
        if rows:
            count = rows[0].get("request_count", 0)
            return f"📊 Today's usage: {count}/{DAILY_LIMIT_GENERATE} generations"
        else:
            return f"📊 Today's usage: 0/{DAILY_LIMIT_GENERATE} generations"

    except Exception:
        return f"📊 Rate limits: {DAILY_LIMIT_GENERATE} generations/day"