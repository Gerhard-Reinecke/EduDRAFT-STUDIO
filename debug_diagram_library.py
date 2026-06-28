# ======================================================================================
# debug_diagram_library.py
# ======================================================================================
# Module: Diagram Pipeline Debugging & Renderer Diagnostic Harness
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module provides a focused debugging harness for the diagram generation pipeline.
#
# It is used to trace the execution of generate_diagram(), inspect returned payloads,
# validate image byte output, and isolate renderer-level failures when specific
# archetypes or rendering paths behave unexpectedly.
#
# It is intended for engineering diagnostics, regression debugging, and direct
# inspection of diagram-library behaviour outside the main application workflow.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Pipeline Invocation Testing
#    - Builds a controlled test request for diagram generation
#    - Calls diagram_library.generate_diagram() directly
#    - Verifies that the public API returns a structured response
#
# 2. Output Contract Inspection
#    - Prints returned keys and response structure
#    - Confirms status, archetype_id, title, message, MIME type, and debug payload
#    - Validates whether returned image bytes are present and non-empty
#
# 3. Renderer-Level Isolation
#    - Detects when a specific archetype has been chosen
#    - Looks up the registered renderer function directly from ARCHETYPES
#    - Executes the renderer in isolation to distinguish:
#        • pipeline-level failure
#        • archetype-resolution failure
#        • renderer implementation failure
#
# 4. Exception Trace Visibility
#    - Captures and prints full stack traces for:
#        • generate_diagram() crashes
#        • direct renderer crashes
#    - Helps pinpoint breakage during development and repair work
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - diagram_library.py → target generation engine under test
# - traceback → detailed exception inspection
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Minimal and direct debugging entry point
# - Developer-readable terminal output
# - Safe isolation of renderer failures
# - Focused regression support for diagram pipeline fixes
#
# --------------------------------------------------------------------------------------
# ROLE IN SYSTEM
# --------------------------------------------------------------------------------------
# This module is a developer diagnostic tool and is NOT part of the production runtime.
#
# It is typically used when:
#   - generate_diagram() returns unexpected output
#   - a specific archetype appears to resolve incorrectly
#   - a renderer may be crashing silently
#   - byte-level PNG validation is needed during diagram debugging
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This module is especially useful after changes to:
#     • ARCHETYPES registry
#     • archetype resolution logic
#     • renderer functions
#     • diagram output contract structure
# - It complements super_doctor.py by providing targeted, manual inspection of one
#   diagram pathway at a time
#
# ======================================================================================

import traceback
import diagram_library as dl

def run_debug():
    req = {
        "prompt": "Year 8 statistics: include a histogram, stem-and-leaf, mean median mode range.",
        "level": "J",
        "subject": "math",
        "archetype_hint": "",
        "params": {}
    }

    print("=== DEBUG: Calling generate_diagram with user_ctx=None ===")
    try:
        out = dl.generate_diagram(req, user_ctx=None)

        print("DEBUG raw out keys:", list(out.keys()) if isinstance(out, dict) else type(out))
        print("DEBUG bytes len:", (len(out.get("bytes")) if isinstance(out, dict) and isinstance(out.get("bytes"), (bytes, bytearray)) else None))
        if out is None:
            print("DEBUG: generate_diagram returned None (BUG).")
            return

        print("DEBUG RESULT status:", out.get("status"))
        print("DEBUG archetype_id:", out.get("archetype_id"))
        print("DEBUG title:", out.get("title"))
        print("DEBUG message:", out.get("message"))          # ✅ THIS is what we need
        print("DEBUG mime:", out.get("mime"))
        print("DEBUG bytes type:", type(out.get("bytes")))
        print("DEBUG debug:", out.get("debug"))

        # If it chose stem_and_leaf_blank, call renderer directly to isolate failure
        if out.get("archetype_id") == "stem_and_leaf_blank":
            print("\n=== DEBUG: Direct renderer call for stem_and_leaf_blank ===")
            meta = dl.ARCHETYPES.get("stem_and_leaf_blank", {})
            renderer_name = meta.get("renderer")
            renderer_fn = getattr(dl, renderer_name, None)
            print("Renderer name:", renderer_name, "callable:", callable(renderer_fn))

            params = (out.get("debug") or {}).get("params") or {}
            print("Params passed:", params)

            try:
                png = renderer_fn(params, level="J", subject="math", title=meta.get("title"))
                print("Direct renderer produced bytes:", isinstance(png, (bytes, bytearray)), "len:", (len(png) if png else None))
            except Exception as e:
                print("DIRECT RENDERER EXCEPTION:", type(e).__name__, str(e))
                traceback.print_exc()

    except Exception as e:
        print("=== DEBUG EXCEPTION (generate_diagram crashed) ===")
        print(type(e).__name__, str(e))
        traceback.print_exc()

if __name__ == "__main__":
    run_debug()