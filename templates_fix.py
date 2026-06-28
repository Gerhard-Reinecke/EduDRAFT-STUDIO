# ======================================================================================
# templates_fix.py
# ======================================================================================
# Module: Template Data Debugging & Supabase Inspection Utility
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module provides a lightweight debugging utility for inspecting template data
# stored in Supabase.
#
# It is used during development and troubleshooting to verify template records,
# diagnose loading issues, and confirm database integrity without affecting the
# main application workflow.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Template Table Inspection
#    - Retrieves all records from the "templates" table
#    - Displays key metadata including:
#        • Template ID
#        • Template name
#        • Associated user
#        • Creation timestamp
#        • Content size (character length)
#
# 2. Data Integrity Debugging
#    - Helps identify:
#        • Missing or malformed template content
#        • Unexpected null values
#        • Duplicate or inconsistent records
#
# 3. Usage Table Visibility
#    - Retrieves and displays records from the "user_usage" table
#    - Provides quick insight into rate limiting and usage tracking data
#
# 4. Console-Based Reporting
#    - Outputs structured debug information to the terminal
#    - Designed for developer visibility rather than UI integration
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - supabase-py client → database access
# - os → environment variable retrieval
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Non-invasive debugging (read-only operations)
# - Minimal dependencies and fast execution
# - Clear, human-readable console output
# - Safe to run independently from the main application
#
# --------------------------------------------------------------------------------------
# ROLE IN SYSTEM
# --------------------------------------------------------------------------------------
# This module is a developer utility and is NOT part of the production runtime.
#
# It is typically used when:
#   - Templates are not loading correctly in the UI
#   - Template content appears corrupted or incomplete
#   - Database state needs verification during development or testing
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - Requires valid SUPABASE_URL and SUPABASE_ANON_KEY environment variables
# - Outputs sensitive metadata (user IDs), so should not be exposed in production logs
# - Intended for local debugging, Colab inspection, or admin diagnostics
#
# ======================================================================================

import os
from supabase import create_client

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip()
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "").strip()
supabase = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)

def debug_templates():
    """Debug template data"""
    print("=== DEBUG TEMPLATES ===")
    
    # Get all templates
    res = supabase.table("templates").select("*").execute()
    rows = getattr(res, "data", None) or []
    
    print(f"Total templates in database: {len(rows)}")
    for i, r in enumerate(rows):
        print(f"\nTemplate #{i+1}:")
        print(f"  ID: {r.get('id')}")
        print(f"  Name: {r.get('name')}")
        print(f"  User: {r.get('user_id', '')[:8]}...")
        print(f"  Created: {r.get('created_at')}")
        print(f"  Content length: {len(r.get('template_md', ''))} chars")
    
    # Also check user_usage table
    print("\n=== USER USAGE ===")
    res2 = supabase.table("user_usage").select("*").execute()
    rows2 = getattr(res2, "data", None) or []
    print(f"User usage records: {len(rows2)}")
    
    return rows

if __name__ == "__main__":
    debug_templates()