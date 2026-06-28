# ======================================================================================
# colab_runner.py
# ======================================================================================
# Module: Execution Orchestrator, Command Router & Backup Control Interface
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module is the primary execution entry point for EduDraft Studio when running
# in Google Colab or similar environments.
#
# It orchestrates system-level operations including diagnostics, backups, and command
# routing, ensuring that the application environment is correctly initialized and that
# critical safety mechanisms (such as backups) are consistently enforced.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Command Routing
#    - Acts as the single command-line interface for system operations
#    - Routes execution to:
#        • system_doctor.py (light diagnostics)
#        • super_doctor.py (full diagnostics)
#        • backup workflows
#
# 2. Environment Initialization
#    - Loads environment variables from .env_local.json (if present)
#    - Ensures required secrets are available before execution
#    - Supports fallback mapping (SUPABASE_KEY → SUPABASE_ANON_KEY)
#
# 3. Automated Backup System
#    - Creates timestamped ZIP backups of the entire project directory
#    - Stores backups in Google Drive under structured folders
#    - Runs automatically before executing commands (unless disabled)
#
# 4. Backup Safety Enforcement
#    - Prevents accidental data loss by enforcing pre-operation backups
#    - Supports optional disabling via NO_BACKUP environment flag
#
# 5. Diagnostic Execution
#    - Executes system-level diagnostic tools in subprocesses:
#        • system_doctor.py → fast checks
#        • super_doctor.py → full system validation
#    - Supports fast-mode diagnostics via environment injection
#
# 6. Process Control
#    - Executes commands in isolated subprocess environments
#    - Returns exit codes for integration with scripts and pipelines
#
# --------------------------------------------------------------------------------------
# SUPPORTED COMMANDS
# --------------------------------------------------------------------------------------
# python colab_runner.py system
#     → Runs lightweight system diagnostics
#
# python colab_runner.py super
#     → Runs full system diagnostics
#
# python colab_runner.py superfast
#     → Runs super_doctor in fast mode (reduced checks)
#
# python colab_runner.py backup [tag]
#     → Creates a manual backup with optional tag
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - os, sys → environment and CLI handling
# - json → local environment loading
# - zipfile, shutil → backup packaging
# - pathlib → filesystem operations
# - datetime → timestamping
# - subprocess → external command execution
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Single entry point for system operations
# - Backup-first safety (no destructive operations without snapshot)
# - Environment-aware execution
# - Minimal assumptions about runtime state
# - Clear command-driven behaviour
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module serves as the operational control interface for developers.
#
# In practice:
#   - It is the first file executed in Colab-based workflows
#   - It ensures the system is safe (backed up) before running diagnostics
#   - It coordinates execution of diagnostic modules
#
# It sits above:
#   - system_doctor.py
#   - super_doctor.py
#   - backup_tools.py (conceptually)
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - Designed specifically for Colab + Google Drive workflows
# - Backup location assumes /content/drive is mounted
# - Backup naming convention is tied to restore logic and should not be changed lightly
# - This module should remain simple, stable, and predictable
#
# ======================================================================================

\
#!/usr/bin/env python3
# colab_runner.py - the only thing you run
# Commands:
#   python colab_runner.py system
#   python colab_runner.py super
#   python colab_runner.py superfast
#   python colab_runner.py backup [tag]
#
# Behavior:
# - Loads secrets from .env_local.json if present
# - ALWAYS creates a Drive backup zip unless NO_BACKUP=1
# - Runs doctors and leaves reports in _doctor_reports

import os, sys, json, zipfile, shutil
from pathlib import Path
from datetime import datetime
import subprocess

PROJECT_ROOT = os.environ.get("PROJECT_ROOT","/content/MarikeApp")

def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def stamp():
    return datetime.now().strftime("%H%M%S")

def load_env_local(root: Path):
    p = root / ".env_local.json"
    if not p.exists():
        return False
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
        for k in ["OPENAI_API_KEY","SUPABASE_URL","SUPABASE_ANON_KEY","SUPABASE_KEY"]:
            if d.get(k) and not os.environ.get(k):
                os.environ[k] = d[k]
        if os.environ.get("SUPABASE_KEY") and not os.environ.get("SUPABASE_ANON_KEY"):
            os.environ["SUPABASE_ANON_KEY"] = os.environ["SUPABASE_KEY"]
        return True
    except Exception:
        return False

def drive_backup(root: Path, tag: str = ""):
    # Requires drive to be mounted in Colab; if not, skip gracefully
    drive = Path("/content/drive")
    if not drive.exists():
        return None

    date_dir = datetime.now().strftime("%Y-%m-%d")
    out_dir = drive / "MyDrive" / "MarikeApp_Backups" / date_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    t = stamp()
    tag_s = f"_{tag}" if tag else ""
    zip_path = out_dir / f"MarikeApp_{t}{tag_s}.zip"

    def should_skip(p: Path):
        name = p.name
        if name in {"__pycache__", ".git"}:
            return True
        if name.endswith(".pyc"):
            return True
        return False

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in root.rglob("*"):
            if should_skip(p):
                continue
            if p.is_dir():
                continue
            rel = p.relative_to(root)
            z.write(p, arcname=str(rel))

    return str(zip_path)

def run(cmd_list):
    return subprocess.run(cmd_list).returncode

def main():
    root = Path(PROJECT_ROOT).resolve()
    os.environ["PROJECT_ROOT"] = str(root)

    if len(sys.argv) < 2:
        print("Usage: python colab_runner.py [system|super|superfast|backup]")
        raise SystemExit(2)

    cmd = sys.argv[1].strip().lower()

    load_env_local(root)

    # Always backup unless disabled
    if os.environ.get("NO_BACKUP","0").strip() != "1":
        tag = cmd
        if cmd == "backup" and len(sys.argv) >= 3:
            tag = sys.argv[2].strip()
        zip_path = drive_backup(root, tag=tag)
        if zip_path:
            print("✅ Backup created:", zip_path)
        else:
            print("WARN: Drive not mounted; backup skipped.")

    if cmd == "system":
        raise SystemExit(run([sys.executable, str(root/"system_doctor.py")]))
    elif cmd == "super":
        raise SystemExit(run([sys.executable, "-u", str(root/"super_doctor.py")]))
    elif cmd == "superfast":
        env = os.environ.copy()
        env["SUPER_DOCTOR_FAST"] = "1"
        raise SystemExit(subprocess.run([sys.executable, "-u", str(root/"super_doctor.py")], env=env).returncode)
    elif cmd == "backup":
        # already done above
        print("OK: backup completed (or skipped).")
        raise SystemExit(0)
    else:
        print("Unknown cmd:", cmd)
        raise SystemExit(2)

if __name__ == "__main__":
    main()
