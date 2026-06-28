# ======================================================================================
# backup_tools.py
# ======================================================================================
# Module: Backup Creation, Versioning & Integrity Verification Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module provides automated backup, versioning, and integrity tracking for the
# EduDraft Studio project.
#
# It creates timestamped project snapshots, detects changes using content hashing,
# generates manifest metadata for each backup, and prevents redundant backups when
# no changes have occurred.
#
# It ensures that system state can be preserved, audited, and restored reliably.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Project Snapshot Creation
#    - Compresses the entire project directory into a ZIP archive
#    - Preserves full directory structure for restoration
#    - Stores backups in date-organised folders
#
# 2. Change Detection (Fingerprinting)
#    - Computes a SHA-256 hash of the project contents
#    - Detects changes at file-content level (not just filenames)
#    - Skips backup creation if no changes are detected (unless forced)
#
# 3. Versioning & Naming
#    - Generates timestamped backup filenames
#    - Supports custom labels for backup identification
#    - Maintains chronological backup history
#
# 4. Manifest Generation
#    - Writes a JSON manifest alongside each backup
#    - Captures:
#        • creation timestamp
#        • project root
#        • backup file path
#        • label
#        • fingerprint hash
#
# 5. Snapshot State Tracking
#    - Stores the most recent fingerprint in a persistent state file
#    - Enables comparison between current and previous project states
#
# 6. Selective File Handling
#    - Excludes unnecessary directories:
#        • __pycache__, .git, virtual environments, caches
#    - Skips backup files themselves to avoid recursive growth
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - os, pathlib → filesystem navigation
# - json → manifest and state tracking
# - hashlib → content fingerprinting
# - zipfile → backup archive creation
# - argparse → CLI interface
# - datetime → timestamping
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Data safety first (no silent overwrites or destructive operations)
# - Content-based change detection (not timestamp-based)
# - Minimal redundancy (skip unchanged backups)
# - Transparent versioning via manifests
# - Portable backup format (ZIP)
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module is the persistent state protection layer of the system.
#
# In practice:
#   - colab_runner.py may trigger backups automatically before operations
#   - Developers can manually create labelled backups via CLI
#   - Backup archives can be used for restore, audit, or rollback
#
# It ensures:
#   - no work is lost
#   - system evolution is trackable
#   - recovery paths always exist
#
# --------------------------------------------------------------------------------------
# CLI USAGE
# --------------------------------------------------------------------------------------
# python backup_tools.py backup [label]
#     → Creates a backup with optional label
#
# python backup_tools.py backup [label] --force
#     → Forces backup even if no changes detected
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - Requires PROJECT_ROOT and BACKUP_BASE environment variables
# - Backup location is typically Google Drive when used in Colab
# - Fingerprint tracking file (.last_fingerprint.json) controls change detection
# - Backup naming convention should remain stable for restore compatibility
#
# ======================================================================================

#!/usr/bin/env python3
import os, json, hashlib, zipfile, pathlib, argparse
from datetime import datetime

PROJECT_ROOT = os.environ.get("PROJECT_ROOT", "/content/MarikeApp")
BACKUP_BASE  = os.environ.get("BACKUP_BASE", "/content/drive/MyDrive/MarikeApp_Backups")
DEFAULT_LABEL = os.environ.get("DEFAULT_LABEL", "work")

SKIP_DIRS = {".git", "__pycache__", ".ipynb_checkpoints", ".cache", "venv", ".venv"}
SKIP_EXTS = {".zip"}  # keep backups out of fingerprint

def _folder_fingerprint(root: str) -> str:
    root = os.path.abspath(root)
    h = hashlib.sha256()
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fn in sorted(filenames):
            if any(fn.endswith(ext) for ext in SKIP_EXTS):
                continue
            if fn.endswith(".bak") or ".bak_" in fn:
                continue
            fp = os.path.join(dirpath, fn)
            rel = os.path.relpath(fp, root).replace("\\", "/")
            h.update(rel.encode("utf-8", errors="ignore"))
            try:
                with open(fp, "rb") as f:
                    while True:
                        chunk = f.read(1024 * 1024)
                        if not chunk:
                            break
                        h.update(chunk)
            except Exception as e:
                h.update(repr(e).encode("utf-8", errors="ignore"))
    return h.hexdigest()

def backup_project(label: str = None, force: bool = False) -> str | None:
    label = (label or DEFAULT_LABEL).strip().replace(" ", "_")
    root = PROJECT_ROOT
    if not os.path.isdir(root):
        raise RuntimeError(f"PROJECT_ROOT not found: {root}")

    out_base = pathlib.Path(BACKUP_BASE)
    out_base.mkdir(parents=True, exist_ok=True)

    date_folder = datetime.now().strftime("%Y-%m-%d")
    out_dir = out_base / date_folder
    out_dir.mkdir(parents=True, exist_ok=True)

    state_file = out_base / ".last_fingerprint.json"
    current_fp = _folder_fingerprint(root)

    last = {}
    if state_file.exists():
        try:
            last = json.loads(state_file.read_text(encoding="utf-8"))
        except:
            last = {}

    if (not force) and last.get("fingerprint") == current_fp:
        print("ℹ️ No changes since last backup — skipping.")
        return None

    stamp = datetime.now().strftime("%H%M%S")
    zip_name = f"MarikeApp_{stamp}_{label}.zip"
    zip_path = out_dir / zip_name

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
            for fn in filenames:
                fp = os.path.join(dirpath, fn)
                rel = os.path.relpath(fp, root).replace("\\", "/")
                z.write(fp, arcname=rel)

    manifest = {
        "created": datetime.now().isoformat(timespec="seconds"),
        "project_root": root,
        "zip": str(zip_path),
        "label": label,
        "fingerprint": current_fp,
    }
    (out_dir / (zip_name + ".manifest.json")).write_text(
        json.dumps(manifest, indent=2), encoding="utf-8"
    )
    state_file.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    print("✅ Backup created:", zip_path)
    return str(zip_path)

def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("backup")
    b.add_argument("label", nargs="?", default="work")
    b.add_argument("--force", action="store_true")
    args = ap.parse_args()
    if args.cmd == "backup":
        backup_project(label=args.label, force=args.force)

if __name__ == "__main__":
    main()
