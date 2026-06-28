# ======================================================================================
# config.py
# ======================================================================================
# Module: Environment Configuration, Service Initialization & Global Settings Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module is the central configuration and initialization layer of EduDraft Studio.
#
# It loads environment variables, initializes external service clients (OpenAI and
# Supabase), defines system-wide constants, and exposes global feature flags and
# operational limits used across all modules.
#
# It ensures that the application starts in a valid, fully-configured state before any
# runtime logic is executed.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Environment Variable Loading
#    - Retrieves required secrets from runtime environment:
#        • OPENAI_API_KEY
#        • SUPABASE_URL
#        • SUPABASE_ANON_KEY
#    - Validates presence of required variables
#    - Raises hard errors if critical configuration is missing
#
# 2. Service Client Initialization
#    - Initializes OpenAI client for LLM and image generation
#    - Initializes Supabase client for authentication, storage, and database access
#    - Exposes shared client instances for system-wide use
#
# 3. Feature Flag Management
#    - Controls optional system features via environment-driven flags:
#        • ENABLE_IMAGE_GEN
#    - Allows runtime toggling of cost-sensitive functionality
#
# 4. Credit System Configuration
#    - Defines cost per image generation
#    - Provides system-wide constants used by billing and diagram modules
#
# 5. Image Generation Defaults
#    - Sets default OpenAI image model and output size
#    - Ensures consistent behaviour across image-generation workflows
#
# 6. Rate Limiting Constants
#    - Defines global request limits for:
#        • document generation
#        • transcription workflows
#
# 7. Template System Configuration
#    - Defines available template categories
#    - Provides controlled vocabulary for UI and storage consistency
#
# 8. Health & Debug Utilities
#    - Provides masked Supabase credential visibility for debugging
#    - Implements lightweight Supabase connectivity check
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - os → environment variable access
# - openai → OpenAI client initialization
# - supabase-py → Supabase client initialization
# - datetime → timestamping utilities (used indirectly by system)
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Fail-fast configuration:
#     Application must not run without required credentials
#
# - Centralised configuration:
#     All global settings defined in one place
#
# - Environment-driven behaviour:
#     No hardcoded secrets or environment assumptions
#
# - Safe debugging:
#     Sensitive values are masked before exposure
#
# - System-wide consistency:
#     Shared constants reduce duplication and drift across modules
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module is the root dependency of the system.
#
# In practice:
#   - llm.py uses the OpenAI client
#   - credits.py and rate_limit.py use Supabase
#   - diagram_library.py uses image-generation flags and cost settings
#   - app.py relies on all configuration values for runtime behaviour
#
# Any failure here prevents the application from starting correctly.
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This module must be loaded successfully before any other module executes
# - Environment variables must be configured in:
#     • Hugging Face Spaces secrets
#     • or local environment for development
# - Feature flags (e.g., ENABLE_IMAGE_GEN) directly affect monetisation behaviour
# - Changes here propagate across the entire system and should be handled with care
#
# ======================================================================================

import os
from datetime import datetime, timezone
from supabase import create_client, Client
from openai import OpenAI

# =============================
# ENV + CLIENTS
# =============================
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip()
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "").strip()

if not OPENAI_API_KEY:
    raise RuntimeError("OPENAI_API_KEY is not set. Add it in your Space Secrets.")

if not SUPABASE_URL or not SUPABASE_ANON_KEY:
    raise RuntimeError("Supabase env vars missing: SUPABASE_URL and/or SUPABASE_ANON_KEY")

client = OpenAI(api_key=OPENAI_API_KEY)
supabase: Client = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)

def _mask(s: str, show: int = 6) -> str:
    if not s:
        return "(missing)"
    s = str(s)
    if len(s) <= show:
        return s
    return s[:show] + "..." + s[-4:]

SUPABASE_DEBUG = (
    "SUPABASE_URL=" + _mask(SUPABASE_URL) + "\n"
    "SUPABASE_ANON_KEY=" + _mask(SUPABASE_ANON_KEY)
)

# =============================
# CREDITS / IMAGE GEN FLAGS (V2.6)
# =============================
ENABLE_IMAGE_GEN = os.environ.get("ENABLE_IMAGE_GEN", "0").strip().lower() in {"1", "true", "yes", "on"}
IMAGE_CREDITS_PER_IMAGE = float(os.environ.get("IMAGE_CREDITS_PER_IMAGE", "1").strip())

# =============================
# IMAGE GEN DEFAULTS (V3.2)
# =============================
OPENAI_IMAGE_MODEL = (os.environ.get("OPENAI_IMAGE_MODEL", "gpt-image-1") or "").strip() or "gpt-image-1"
OPENAI_IMAGE_SIZE  = (os.environ.get("OPENAI_IMAGE_SIZE",  "1024x1024")  or "").strip() or "1024x1024"

# =============================
# RATE LIMITING
# =============================
DAILY_LIMIT_GENERATE = 50
DAILY_LIMIT_TRANSCRIBE = 100

# =============================
# TEMPLATE SYSTEM
# =============================
TEMPLATE_CATEGORIES = [
    "Test/Quiz", "Worksheet", "Homework", "Exam", "Lesson Plan",
    "Investigation", "Rubric", "Marking Key", "Revision Sheet", "Custom"
]

def supabase_secrets_masked() -> str:
    return SUPABASE_DEBUG

def supabase_healthcheck():
    try:
        supabase.table("drafts").select("id").limit(1).execute()
        return "✅ Supabase connection OK"
    except Exception as e:
        return f"❌ Supabase error: {type(e).__name__}: {e}"
