# ======================================================================================
# system_doctor.py
# ======================================================================================
# Module: Lightweight System Integrity & Environment Sanity Check
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module performs a minimal, fast, and safe diagnostic check of the core system
# environment and file structure.
#
# It verifies the presence of required modules, lists available Python files, and
# confirms the availability of critical environment variables (e.g., OpenAI and
# Supabase configuration).
#
# It is designed as a quick health-check tool to validate that the system is in a
# runnable state before deeper diagnostics or execution.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Required File Validation
#    - Checks for the presence of essential system modules
#    - Identifies missing critical files that would break the application
#
# 2. Project Structure Snapshot
#    - Lists all Python files in the project root
#    - Provides a quick overview of the current codebase state
#
# 3. Environment Variable Verification
#    - Confirms presence of required runtime variables:
#        • OPENAI_API_KEY
#        • SUPABASE_URL
#        • SUPABASE_ANON_KEY
#    - Does NOT expose sensitive values (boolean presence only)
#
# 4. Reporting
#    - Writes structured JSON report
#    - Writes human-readable text report
#    - Stores outputs in the _doctor_reports directory
#
# --------------------------------------------------------------------------------------
# OUTPUTS
# --------------------------------------------------------------------------------------
# Reports:
#   - _doctor_reports/system_doctor.json
#   - _doctor_reports/system_doctor.txt
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - Standard library:
#     os, json, hashlib, pathlib, datetime
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Fast execution (no heavy checks or external calls)
# - Safe diagnostics (no imports or code execution)
# - Zero side effects beyond report generation
# - ASCII-safe output for maximum compatibility
#
# --------------------------------------------------------------------------------------
# ROLE IN SYSTEM
# --------------------------------------------------------------------------------------
# This module is the "quick check" companion to super_doctor.py:
#
# - system_doctor.py:
#     • Fast
#     • Minimal
#     • Safe to run anytime
#
# - super_doctor.py:
#     • Deep diagnostics
#     • Full system validation
#     • Slower but comprehensive
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - Intended for rapid debugging and environment validation
# - Useful before deployment or after environment changes
# - Does NOT validate syntax, imports, or runtime behaviour
# - Can be safely executed in restricted environments (e.g., Colab, CI pipelines)
#
# ======================================================================================

#!/usr/bin/env python3
# system_doctor.py - minimal safe diagnostics (ASCII only)

import os, json, hashlib
from pathlib import Path
from datetime import datetime

REQUIRED = [
  "app.py","exports.py","diagram_library.py","auth.py","config.py","llm.py",
  "rate_limit.py","credits.py","ingest_docx.py","ingest_pdf.py","ingest_pptx.py",
  "app_behavior_doctor.py"
]

def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1024*1024), b""):
            h.update(chunk)
    return h.hexdigest()

def main():
    root = Path(os.environ.get("PROJECT_ROOT","/content/MarikeApp")).resolve()
    os.environ["PROJECT_ROOT"] = str(root)
    rep = root / "_doctor_reports"
    rep.mkdir(parents=True, exist_ok=True)

    miss = [x for x in REQUIRED if not (root/x).exists()]
    py_files = sorted([p.name for p in root.glob("*.py") if p.is_file()])

    report = {
        "time": datetime.now().isoformat(timespec="seconds"),
        "project_root": str(root),
        "missing_required": miss,
        "py_files": py_files,
        "env": {
            "OPENAI_API_KEY": bool(os.environ.get("OPENAI_API_KEY")),
            "SUPABASE_URL": bool(os.environ.get("SUPABASE_URL")),
            "SUPABASE_ANON_KEY": bool(os.environ.get("SUPABASE_ANON_KEY")),
        }
    }

    (rep/"system_doctor.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    (rep/"system_doctor.txt").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print("OK: wrote reports to", rep)

if __name__ == "__main__":
    main()
