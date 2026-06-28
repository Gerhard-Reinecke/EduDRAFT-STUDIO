# ======================================================================================
# app_behavior_doctor.py
# ======================================================================================
# Module: End-to-End Diagram Pipeline Validation & Behaviour Diagnostic Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module performs full end-to-end validation of the diagram generation pipeline.
#
# It verifies that diagram requests flow correctly from VISUAL placeholders through
# exports.py into diagram_library.py, and that valid image outputs are produced for
# each registered archetype.
#
# It is designed as a behavioural validation tool to confirm that the system works
# as intended under real execution conditions, not just at module level.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Archetype Enumeration
#    - Loads all registered diagram archetypes from diagram_library.ARCHETYPES
#    - Iterates through each archetype for validation testing
#
# 2. End-to-End Pipeline Execution
#    - Constructs synthetic [[VISUAL ...]] placeholders for each archetype
#    - Routes them through exports.extract_visuals_from_line(...)
#    - Triggers the full rendering pipeline (including diagram generation)
#
# 3. Output Validation
#    - Verifies that image files are created for each archetype
#    - Checks file size against minimum thresholds
#    - Validates PNG integrity using PIL
#
# 4. Integration Testing
#    - Confirms correct interaction between:
#        • exports.py (visual extraction & routing)
#        • diagram_library.py (diagram generation engine)
#    - Detects failures at the integration level rather than isolated modules
#
# 5. Failure Diagnostics
#    - Captures exceptions and stack traces for failed archetypes
#    - Reports detailed error messages and failure causes
#    - Provides targeted rerun instructions using ONLY_KEYS filtering
#
# 6. Reporting & Summary
#    - Outputs pass/fail status for each archetype
#    - Provides aggregated summary of results
#    - Lists failing cases with debug details
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - diagram_library.py → archetype registry and rendering engine
# - exports.py → VISUAL parsing and pipeline routing
# - PIL (Pillow) → image validation
# - Standard library → filesystem, timing, and subprocess utilities
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - End-to-end validation over unit testing
# - Deterministic reproducibility of test cases
# - Clear pass/fail output for rapid diagnosis
# - Minimal assumptions about upstream state
# - Fast isolation of failing archetypes
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module is a behavioural test harness for the diagram subsystem.
#
# In practice:
#   - super_doctor.py checks system-level health
#   - debug_diagram_library.py isolates individual diagram issues
#   - app_behavior_doctor.py validates the FULL pipeline under real conditions
#
# It ensures:
#   - diagram rendering works in production context
#   - export integration is functioning correctly
#   - outputs are valid and usable by downstream systems
#
# --------------------------------------------------------------------------------------
# RUNTIME CONFIGURATION
# --------------------------------------------------------------------------------------
# Environment variables:
#
# - PROJECT_ROOT
#     → root directory of the project
#
# - ONLY_KEYS="k1,k2,..."
#     → restrict testing to specific archetypes
#
# - MIN_PNG_BYTES
#     → minimum acceptable file size for output validation
#
# - ENABLE_IMAGE_GEN
#     → affects fallback behaviour for non-deterministic diagrams
#
# - KEEP_OUTPUTS=1
#     → preserves output directory between runs
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This module is critical for regression testing after:
#     • archetype additions
#     • renderer changes
#     • exports pipeline updates
#
# - Failures here indicate real pipeline issues, not theoretical ones
# - This is the closest simulation of production diagram behaviour
# - Should be run before deployment or major releases
#
# ======================================================================================

#!/usr/bin/env python3
"""
EduDraft Studio - Diagram System Validation & Diagnostics Tool

Purpose:
Validates end-to-end diagram generation across all registered archetypes.
Ensures integration between diagram_library and exports modules, including:
- Archetype enumeration
- Visual block construction
- Image generation pipeline
- Output validation (existence, size, integrity)

This tool is used for:
- System verification
- Regression testing
- Pipeline debugging
- Production readiness checks

Run:
  ONLY_KEYS="k1,k2" PROJECT_ROOT="/path" python app_behavior_doctor.py
"""

from __future__ import annotations
import os, sys, traceback, time, shutil
from pathlib import Path

def _now():
    return time.strftime("%Y-%m-%d %H:%M:%S")

def _h(title: str):
    print("\n" + "=" * 78)
    print(title)
    print("=" * 78)

def _ok(s: str): print("OK  :", s)
def _bad(s: str): print("FAIL:", s)
def _warn(s: str): print("WARN:", s)

def _find_root() -> Path:
    env = os.environ.get("PROJECT_ROOT", "").strip()
    if env:
        return Path(env).expanduser().resolve()
    return Path.cwd().resolve()

def _ensure_clean_dir(p: Path):
    keep = os.environ.get("KEEP_OUTPUTS","0") == "1"
    if p.exists() and not keep:
        shutil.rmtree(p, ignore_errors=True)
    p.mkdir(parents=True, exist_ok=True)

def _png_ok(path: Path, min_bytes: int) -> tuple[bool, str]:
    if not path.exists():
        return False, "PNG file not created"
    size = path.stat().st_size
    if size < min_bytes:
        return False, f"PNG too small ({size} bytes < {min_bytes})"
    try:
        from PIL import Image
        with Image.open(path) as im:
            im.verify()
        return True, f"PNG OK ({size} bytes)"
    except Exception as e:
        return False, f"PIL verify failed: {type(e).__name__}: {e}"

def main():
    root = _find_root()
    sys.path.insert(0, str(root))

    _h("APP BEHAVIOR DOCTOR")
    print("Time:", _now())
    print("PROJECT_ROOT:", root)

    try:
        import diagram_library
        import exports
    except Exception as e:
        _bad(f"Import failed: {type(e).__name__}: {e}")
        print(traceback.format_exc(limit=8))
        return 2

    ARCH = getattr(diagram_library, "ARCHETYPES", None)
    if not isinstance(ARCH, dict) or not ARCH:
        _bad("diagram_library.ARCHETYPES not found or empty.")
        return 2

    keys = list(ARCH.keys())

    only = os.environ.get("ONLY_KEYS","").strip()
    if only:
        wanted = [k.strip() for k in only.split(",") if k.strip()]
        keys = [k for k in keys if k in wanted]
        _warn("ONLY_KEYS active: " + ", ".join(keys))

    min_bytes = int(os.environ.get("MIN_PNG_BYTES","8000"))
    out_dir = root / "_doctor_outputs"
    _ensure_clean_dir(out_dir)

    if not hasattr(exports, "extract_visuals_from_line"):
        _bad("exports.extract_visuals_from_line not found.")
        return 2

    extract_visuals_from_line = exports.extract_visuals_from_line

    _h("RUN")
    print("Archetypes:", len(keys))
    print("Output dir:", out_dir)
    print("MIN_PNG_BYTES:", min_bytes)
    print("ENABLE_IMAGE_GEN:", os.environ.get("ENABLE_IMAGE_GEN","(unset)"))

    passed = []
    failed = []

    for i, key in enumerate(keys, 1):
        vid = f"d{i:03d}"
        where = f"Doctor test: {key}"
        prompt = f"Deterministic diagram for {key}. Use default params."

        line = f'[[VISUAL id="{vid}" kind="diagram" subtype="{key}" where="{where}" prompt="{prompt}"]]'

        try:
            res = extract_visuals_from_line(line)

            # Expected: (cleaned_text, [image_paths])
            if isinstance(res, tuple) and len(res) == 2:
                cleaned, imgs = res
            else:
                raise RuntimeError(f"Unexpected return type from extract_visuals_from_line: {type(res)}")

            if not isinstance(imgs, list):
                raise RuntimeError(f"Expected list of image paths, got: {type(imgs)}")

            if not imgs:
                raise RuntimeError("No images were produced by extract_visuals_from_line (imgs list is empty)")

            # Validate each produced image
            all_ok = True
            msgs = []
            for p in imgs:
                pth = Path(p)
                ok, msg = _png_ok(pth, min_bytes)
                msgs.append(f"{pth.name}: {msg}")
                if not ok:
                    all_ok = False

            if all_ok:
                _ok(f"{key} -> " + "; ".join(msgs))
                passed.append(key)
            else:
                _bad(f"{key} -> " + "; ".join(msgs))
                failed.append((key, "; ".join(msgs), ""))

        except Exception as e:
            tb = traceback.format_exc(limit=10)
            _bad(f"{key} -> {type(e).__name__}: {e}")
            failed.append((key, f"{type(e).__name__}: {e}", tb))

    _h("SUMMARY")
    print("PASS:", len(passed))
    print("FAIL:", len(failed))

    if failed:
        _h("FAIL DETAILS (first 10)")
        for key, msg, tb in failed[:10]:
            print("-" * 78)
            print("KEY:", key)
            print("ERROR:", msg)
            if tb:
                print(tb)

        print("\nTip: rerun only failures with:")
        print('  ONLY_KEYS="{}" PROJECT_ROOT="{}" python app_behavior_doctor.py'.format(
            ",".join([k for k,_,_ in failed[:10]]), str(root)
        ))

    _h("DONE")
    return 0 if not failed else 1

if __name__ == "__main__":
    raise SystemExit(main())
