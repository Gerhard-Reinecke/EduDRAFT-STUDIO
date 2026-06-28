# ======================================================================================
# super_doctor.py
# ======================================================================================
# Module: System Diagnostics, Integrity Audit & Smoke Test Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module performs comprehensive system-level diagnostics across the EduDraft Studio
# codebase.
#
# It verifies file integrity, checks Python syntax, tests import safety in isolated
# subprocesses, runs diagram rendering diagnostics across supported archetypes, performs
# export smoke tests, and validates ingestion pipelines where sample files are present.
#
# Its role is to detect breakage early, surface unstable modules, and provide a reliable
# health snapshot of the application before or after major development changes.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. File Integrity Snapshot & Change Detection
#    - Computes SHA-256 hashes for key project files
#    - Compares current hashes against prior snapshots
#    - Detects changed or missing files across runs
#
# 2. Syntax Verification
#    - Compiles all Python files in the project root
#    - Reports syntax failures without executing application logic
#
# 3. Safe Import Diagnostics
#    - Tests core module imports in isolated subprocesses
#    - Applies timeouts to prevent hanging imports
#    - Forces local project resolution via PYTHONPATH injection
#
# 4. Diagram Rendering Diagnostics
#    - Iterates across registered diagram archetypes
#    - Builds safe sample parameter sets for each renderer
#    - Executes diagram generation and records successes, fallbacks, and failures
#    - Writes rendered outputs for inspection
#
# 5. Export Smoke Testing
#    - Tests PPTX export generation
#    - Tests DOCX export generation when Pandoc is available
#    - Confirms that export pathways remain operational
#
# 6. Ingestion Smoke Testing
#    - Tests DOCX, PPTX, and PDF extraction pipelines when sample files exist
#    - Reports extracted character counts as a basic ingestion health check
#
# 7. Reporting & Audit Output
#    - Writes structured JSON diagnostics report
#    - Writes human-readable text report
#    - Stores generated diagnostic outputs for manual review
#
# --------------------------------------------------------------------------------------
# OUTPUTS
# --------------------------------------------------------------------------------------
# Reports:
#   - _doctor_reports/super_doctor.json
#   - _doctor_reports/super_doctor.txt
#
# Diagnostic outputs:
#   - _doctor_outputs/diagrams/*.png
#   - _doctor_outputs/exports_smoke/*
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - Standard library:
#     os, sys, json, hashlib, subprocess, py_compile, shutil, pathlib, datetime
# - Local project modules:
#     config, auth, llm, rate_limit, credits, exports, diagram_library,
#     ingest_docx, ingest_pdf, ingest_pptx
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Fail-safe diagnostics (test without breaking the app)
# - Isolated import checking (subprocess + timeout protection)
# - Broad system coverage from a single execution point
# - Repeatable integrity auditing through persistent snapshots
# - Actionable reporting for fast debugging and regression detection
#
# --------------------------------------------------------------------------------------
# OPERATIONAL MODES
# --------------------------------------------------------------------------------------
# - Standard mode:
#     Runs full diagnostics including exports and ingestion smoke tests
#
# - Fast mode:
#     Enabled via SUPER_DOCTOR_FAST=1
#     Skips slower export and ingestion checks for rapid health verification
#
# - Filtered diagram mode:
#     Enabled via ONLY_KEYS
#     Restricts diagram diagnostics to selected archetypes only
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This module is intended for engineering verification, not end-user workflows
# - Import testing is deliberately isolated to prevent blocking or hidden side effects
# - Sample parameter generation is defensive and designed to satisfy renderer shape requirements
# - This is a critical regression-detection tool after major upgrades or refactors
#
# ======================================================================================

#!/usr/bin/env python3
# super_doctor.py
# Smart diagnostics:
# - file integrity snapshot + diff
# - syntax compile all .py
# - import checks (SUBPROCESS + TIMEOUT, so it cannot hang)
# - diagram doctor across ARCHETYPES (ONLY_KEYS supported)
# - export smoke (PPTX always; DOCX if pandoc exists)
# - ingest smoke (if sample files exist)
# Writes:
#   _doctor_reports/super_doctor.json
#   _doctor_reports/super_doctor.txt
# Outputs:
#   _doctor_outputs/diagrams/*.png
#   _doctor_outputs/exports_smoke/*

import os, sys, json, hashlib, subprocess, py_compile, shutil
from pathlib import Path
from datetime import datetime

MODULES = ["config","auth","llm","rate_limit","credits","exports","diagram_library","ingest_docx","ingest_pdf","ingest_pptx"]

def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024*1024), b""):
            h.update(chunk)
    return h.hexdigest()

def which(cmd: str) -> bool:
    return bool(shutil.which(cmd))

def load_env_local(root: Path):
    # subprocess-safe secrets bridge: root/.env_local.json
    p = root / ".env_local.json"
    if not p.exists():
        return False
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
        for k in ["OPENAI_API_KEY","SUPABASE_URL","SUPABASE_ANON_KEY","SUPABASE_KEY"]:
            if d.get(k) and not os.environ.get(k):
                os.environ[k] = d[k]
        # alias
        if os.environ.get("SUPABASE_KEY") and not os.environ.get("SUPABASE_ANON_KEY"):
            os.environ["SUPABASE_ANON_KEY"] = os.environ["SUPABASE_KEY"]
        return True
    except Exception:
        return False

def safe_import(modname: str, timeout_sec: int = 15):
    """
    Import in a separate process so we cannot hang the doctor.
    Also forces PYTHONPATH to include PROJECT_ROOT so it imports *your* local modules.
    """
    cmd = [sys.executable, "-c", f"import {modname}"]
    env = os.environ.copy()
    env["DOCTOR_MODE"] = "1"

    root = Path(os.environ.get("PROJECT_ROOT", "/content/MarikeApp")).resolve()
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = str(root) + (os.pathsep + existing if existing else "")

    # Some imports can be slow on Colab (python-pptx etc.)
    # Give ingestion modules a little more breathing room.
    if modname in {"ingest_pptx","ingest_docx","ingest_pdf"}:
        timeout_sec = max(timeout_sec, 30)

    try:
        r = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=timeout_sec)
        if r.returncode == 0:
            return True, None
        err = (r.stderr or r.stdout or "").strip()
        if not err:
            err = f"Import failed with code {r.returncode}"
        return False, err[:900]
    except subprocess.TimeoutExpired:
        return False, f"TIMEOUT after {timeout_sec}s (module import may be blocking at import-time)"

def _sample_params_for_required(required_params):
    """
    Build safe sample params based on required param names.
    Goal: ALWAYS provide the right *shape* (non-empty lists / dicts) so renderers don't fail.
    """
    rp = list(required_params or [])
    params = {}

    # handy defaults we reuse
    xy_points = [
        {"x": 1, "y": 2, "label": "A"},
        {"x": 2, "y": 1, "label": "B"},
        {"x": 3, "y": 4, "label": "C"},
    ]
    x_list = [1, 2, 3, 4]
    y_list = [2, 1, 4, 3]

    # circuits: keep to types your renderers accept
    circuit_components = [
        {"type": "lamp", "label": "L1"},
        {"type": "resistor", "label": "R1"},
    ]

    for k in rp:
        lk = str(k).lower().strip()

        # -------------------------
        # GRAPHS / POINTS
        # -------------------------
        if lk in {"points", "xy", "pairs"}:
            # works for coordinate_plane_points + scatter_plot style renderers
            params[k] = list(xy_points)

        elif lk in {"x", "x_values", "xvals", "xs"}:
            # line_graph/scatter_plot often want list-like x
            params[k] = list(x_list)

        elif lk in {"y", "y_values", "yvals", "ys"}:
            # line_graph/scatter_plot often want list-like y
            params[k] = list(y_list)

        elif lk in {"p1", "p2"}:
            # slope_triangle / gradient_rise_run style
            params[k] = {"x": 0, "y": 0} if lk == "p1" else {"x": 4, "y": 2}

        # -------------------------
        # BOX & WHISKER
        # -------------------------
        elif lk in {"five_number_summary", "five_num", "summary"}:
            params[k] = {"min": 2, "q1": 4, "median": 6, "q3": 8, "max": 10}

        # Some versions use these directly:
        elif lk in {"min", "minimum"}:
            params[k] = 0
        elif lk in {"max", "maximum"}:
            params[k] = 10
        elif lk in {"q1", "lower_quartile"}:
            params[k] = 3
        elif lk in {"median"}:
            params[k] = 5
        elif lk in {"q3", "upper_quartile"}:
            params[k] = 7

        # -------------------------
        # FOOD WEB
        # -------------------------
        elif lk == "nodes":
            params[k] = [
                {"id": "grass", "label": "Grass"},
                {"id": "rabbit", "label": "Rabbit"},
                {"id": "fox", "label": "Fox"},
            ]
        elif lk in {"links", "edges"}:
            params[k] = [
                {"from": "grass", "to": "rabbit"},
                {"from": "rabbit", "to": "fox"},
            ]

        # -------------------------
        # FREE BODY DIAGRAM
        # -------------------------
        elif lk == "forces":
            params[k] = [
                {"label": "Weight", "angle_deg": 270, "magnitude": 10},
                {"label": "Normal", "angle_deg": 90, "magnitude": 10},
                {"label": "Push", "angle_deg": 0, "magnitude": 6},
            ]

        # -------------------------
        # NUMBER LINE
        # -------------------------
        elif lk in {"start", "x0", "left"}:
            params[k] = 0
        elif lk in {"end", "x1", "right"}:
            params[k] = 10
        elif lk in {"tick_step", "step", "interval"}:
            params[k] = 1

        # -------------------------
        # PROBABILITY TREE (exactly 2 stages)
        # -------------------------
        elif lk == "stages":
            params[k] = [
                {
                    "label": "Stage 1",
                    "branches": [{"label": "A", "p": 0.5}, {"label": "B", "p": 0.5}],
                },
                {
                    "label": "Stage 2",
                    "branches": [{"label": "Yes", "p": 0.5}, {"label": "No", "p": 0.5}],
                },
            ]

        # -------------------------
        # TWO-WAY TABLE
        # -------------------------
        elif lk in {"row_labels", "rowlabels"}:
            params[k] = ["Row 1", "Row 2"]
        elif lk in {"col_labels", "collabels", "column_labels"}:
            params[k] = ["Col 1", "Col 2"]
        elif lk in {"values_matrix", "matrix"}:
            params[k] = [
                [1, 2],
                [3, 4],
            ]
        
        # -------------------------
        # BASIC SERIES DATA
        # -------------------------
        elif lk in {"categories", "labels"}:
            params[k] = ["A", "B", "C"]

        elif lk in {"values", "frequencies"}:
            params[k] = [2, 5, 3]

        elif lk in {"data", "dataset", "numbers"}:
            params[k] = [12, 14, 15, 17, 19, 20, 21]

        elif lk == "bins":
            params[k] = [0, 10, 20, 30]  # histogram convention

        # -------------------------
        # FREQUENCY TABLE (your error demands list of dicts)
        # -------------------------
        elif lk == "rows":
            params[k] = [
                {"label": "A", "freq": 2},
                {"label": "B", "freq": 5},
                {"label": "C", "freq": 3},
            ]

        # -------------------------
        # CIRCUITS (your errors demand non-empty lists)
        # -------------------------
        elif lk == "mode":
            params[k] = "series"  # safe default; valid for series/parallel selectors

        elif lk == "components":
            # circuit_series_parallel expects list of dict components
            params[k] = list(circuit_components)

        elif lk == "loads":
            # series_circuit expects non-empty list (use dicts to be safe)
            params[k] = list(circuit_components)

        elif lk == "branches":
            # parallel_circuit expects list of branches, each branch a list of loads
            params[k] = [
                [circuit_components[0]],
                [circuit_components[0], circuit_components[1]],
            ]

        # -------------------------
        # DEFAULT SAFE FALLBACK
        # -------------------------
        else:
            # non-empty and harmless
            params[k] = "A"

    return params

def main():
    root = Path(os.environ.get("PROJECT_ROOT","/content/MarikeApp")).resolve()
    os.environ["PROJECT_ROOT"] = str(root)

    # secrets bridge
    load_env_local(root)

    # ensure local imports resolve to this project
    sys.path.insert(0, str(root))

    rep_dir = root / "_doctor_reports"
    out_dir = root / "_doctor_outputs"
    rep_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    FAST = (os.environ.get("SUPER_DOCTOR_FAST","0").strip() == "1")

    report = {
        "time": ts(),
        "project_root": str(root),
        "env": {
            "OPENAI_API_KEY": bool(os.environ.get("OPENAI_API_KEY")),
            "SUPABASE_URL": bool(os.environ.get("SUPABASE_URL")),
            "SUPABASE_ANON_KEY": bool(os.environ.get("SUPABASE_ANON_KEY")),
        },
        "sections": {}
    }

    print("[SUPER_DOCTOR] 1/6 file integrity...", flush=True)
    key_files = [
        "app.py","exports.py","diagram_library.py",
        "ingest_docx.py","ingest_pdf.py","ingest_pptx.py",
        "auth.py","config.py","credits.py","llm.py","rate_limit.py",
        "app_behavior_doctor.py","debug_diagram_library.py","system_doctor.py",
        "colab_runner.py","super_doctor.py"
    ]
    hashes = {}
    missing = []
    for fn in key_files:
        fp = root / fn
        if fp.exists():
            hashes[fn] = sha256_file(fp)
        else:
            missing.append(fn)

    snap_path = rep_dir / "file_hashes.json"
    prev = None
    if snap_path.exists():
        try:
            prev = json.loads(snap_path.read_text(encoding="utf-8"))
        except Exception:
            prev = None

    changed = []
    if prev and isinstance(prev, dict) and "hashes" in prev:
        old = prev["hashes"]
        for fn, h in hashes.items():
            if fn in old and old[fn] != h:
                changed.append(fn)

    snap = {"time": ts(), "hashes": hashes, "missing": missing}
    snap_path.write_text(json.dumps(snap, indent=2), encoding="utf-8")
    report["sections"]["file_integrity"] = {"missing": missing, "changed_since_last": changed, "snapshot_path": str(snap_path)}

    print("[SUPER_DOCTOR] 2/6 syntax check...", flush=True)
    py_files = sorted([p for p in root.glob("*.py") if p.is_file()])
    syntax_ok = 0
    syntax_fail = []
    for pth in py_files:
        try:
            py_compile.compile(str(pth), doraise=True)
            syntax_ok += 1
        except Exception as e:
            syntax_fail.append({"file": pth.name, "error": f"{type(e).__name__}: {e}"})
    report["sections"]["syntax"] = {"ok": syntax_ok, "fail": len(syntax_fail), "fails": syntax_fail}

    print("[SUPER_DOCTOR] 3/6 imports (subprocess+timeout)...", flush=True)
    imports_ok = []
    imports_fail = {}
    for m in MODULES:
        ok, err = safe_import(m, timeout_sec=15)
        if ok:
            imports_ok.append(m)
        else:
            imports_fail[m] = err
    report["sections"]["imports"] = {"ok": imports_ok, "fail": imports_fail}

    print("[SUPER_DOCTOR] 4/6 diagram doctor...", flush=True)
    diagrams = {"ok": 0, "fallback": 0, "error": 0, "saved_dir": str(out_dir / "diagrams"), "top_errors": []}
    diag_out = out_dir / "diagrams"
    diag_out.mkdir(parents=True, exist_ok=True)

    only = os.environ.get("ONLY_KEYS","").strip()
    only_set = set([x.strip() for x in only.split(",") if x.strip()]) if only else None

    try:
        import diagram_library as dl
        keys = sorted(getattr(dl, "ARCHETYPES", {}).keys())
        if only_set is not None:
            keys = [k for k in keys if k in only_set]

        for k in keys:
            meta = dl.ARCHETYPES.get(k, {}) or {}
            title = meta.get("title") or k
            req_params = _sample_params_for_required(meta.get("required_params", []))
            req = {
                "where": title,
                "prompt": title,
                "notes": "",
                "subject": meta.get("subject",""),
                "level": "J",
                "archetype_hint": k,
                "params": req_params,
                "auto_image_gen": False,
            }
            try:
                res = dl.generate_diagram(req, user_ctx=None)
                st = (res.get("status") or "").lower()
                if st == "ok":
                    b = res.get("bytes") or b""
                    if isinstance(b, (bytes, bytearray)) and len(b) > 0:
                        (diag_out / f"{k}.png").write_bytes(bytes(b))
                        diagrams["ok"] += 1
                    else:
                        diagrams["error"] += 1
                        if len(diagrams["top_errors"]) < 25:
                            diagrams["top_errors"].append({"key": k, "error": "status ok but bytes missing"})
                elif st == "fallback":
                    diagrams["fallback"] += 1
                else:
                    diagrams["error"] += 1
                    if len(diagrams["top_errors"]) < 25:
                        diagrams["top_errors"].append({"key": k, "error": res.get("message") or res.get("reason") or "error"})
            except Exception as e:
                diagrams["error"] += 1
                if len(diagrams["top_errors"]) < 25:
                    diagrams["top_errors"].append({"key": k, "error": f"{type(e).__name__}: {e}"})
    except Exception as e:
        diagrams["error"] += 1
        diagrams["top_errors"] = [{"key":"diagram_doctor", "error": f"{type(e).__name__}: {e}"}]

    report["sections"]["diagram_doctor"] = diagrams

    # 5) export smoke
    if FAST:
        report["sections"]["export_smoke"] = {"skipped": True, "reason": "SUPER_DOCTOR_FAST=1"}
    else:
        print("[SUPER_DOCTOR] 5/6 export smoke...", flush=True)
        exports_section = {"pptx": None, "docx": None}
        try:
            import exports as ex
            import diagram_library as dl

            smoke_dir = out_dir / "exports_smoke"
            smoke_dir.mkdir(parents=True, exist_ok=True)

            keys = sorted(getattr(dl, "ARCHETYPES", {}).keys())
            pick = keys[:3] if len(keys) >= 3 else keys

            outline_lines = ["Slide 1: Export Smoke", "- This should contain visuals if rendering works."]
            for i, kk in enumerate(pick, 1):
                outline_lines.append(f'- [[VISUAL id="sp{i}" kind="diagram" subtype="{kk}" where="PPT export" prompt="Render archetype {kk}" notes=""]]')
            outline = "\n".join(outline_lines)

            pptx_path = str(smoke_dir / "export_smoke.pptx")
            ex.outline_to_pptx_with_math(outline, pptx_path)
            exports_section["pptx"] = {"path": pptx_path, "exists": Path(pptx_path).exists()}

            docx_path = str(smoke_dir / "export_smoke.docx")
            if which("pandoc"):
                md_lines = ["# Export Smoke Test", ""]
                for i, kk in enumerate(pick, 1):
                    md_lines.append(f'[[VISUAL id="sd{i}" kind="diagram" subtype="{kk}" where="Doctor export" prompt="Render archetype {kk}" notes=""]]')
                    md_lines.append("")
                md_text = "\n".join(md_lines)
                ex.md_to_docx_with_editable_equations(md_text, docx_path, pre_rendered=False)
                exports_section["docx"] = {"path": docx_path, "exists": Path(docx_path).exists()}
            else:
                exports_section["docx"] = {"skipped": True, "reason": "pandoc not installed"}
        except Exception as e:
            exports_section["error"] = f"{type(e).__name__}: {e}"
        report["sections"]["export_smoke"] = exports_section

    # 6) ingest smoke
    if FAST:
        report["sections"]["ingest_smoke"] = {"skipped": True, "reason": "SUPER_DOCTOR_FAST=1"}
    else:
        print("[SUPER_DOCTOR] 6/6 ingest smoke...", flush=True)
        ingest_section = {"docx": None, "pptx": None, "pdf": None}
        try:
            import ingest_docx, ingest_pptx, ingest_pdf
            smoke_docx = root / "_SMOKE_TEST.docx"
            smoke_pptx = root / "_SMOKE_TEST.pptx"

            if smoke_docx.exists() and hasattr(ingest_docx, "extract_text_from_docx"):
                txt = ingest_docx.extract_text_from_docx(str(smoke_docx))
                ingest_section["docx"] = {"path": str(smoke_docx), "chars": len(txt)}
            else:
                ingest_section["docx"] = {"skipped": True, "reason": "_SMOKE_TEST.docx missing or extractor not found"}

            if smoke_pptx.exists() and hasattr(ingest_pptx, "extract_text_from_pptx"):
                txt = ingest_pptx.extract_text_from_pptx(str(smoke_pptx))
                ingest_section["pptx"] = {"path": str(smoke_pptx), "chars": len(txt)}
            else:
                ingest_section["pptx"] = {"skipped": True, "reason": "_SMOKE_TEST.pptx missing or extractor not found"}

            pdfs = list(root.glob("*.pdf"))
            if pdfs and hasattr(ingest_pdf, "extract_text_from_pdf"):
                txt = ingest_pdf.extract_text_from_pdf(str(pdfs[0]))
                ingest_section["pdf"] = {"path": str(pdfs[0]), "chars": len(txt)}
            else:
                ingest_section["pdf"] = {"skipped": True, "reason": "no sample pdf or extractor not found"}
        except Exception as e:
            ingest_section["error"] = f"{type(e).__name__}: {e}"
        report["sections"]["ingest_smoke"] = ingest_section

    print("[SUPER_DOCTOR] writing reports...", flush=True)
    json_path = rep_dir / "super_doctor.json"
    txt_path  = rep_dir / "super_doctor.txt"
    json_path.write_text(json.dumps(report, indent=2), encoding="utf-8")

    lines = []
    lines.append("="*78)
    lines.append("SUPER DOCTOR")
    lines.append("="*78)
    lines.append(f"Time: {report['time']}")
    lines.append(f"PROJECT_ROOT: {report['project_root']}")
    lines.append("")
    lines.append("ENV:")
    for k,v in report["env"].items():
        lines.append(f" - {k}: {v}")
    lines.append("")
    fi = report["sections"]["file_integrity"]
    lines.append("FILE INTEGRITY:")
    lines.append(f" - Missing: {len(fi['missing'])} -> {fi['missing'][:20]}")
    lines.append(f" - Changed since last: {len(fi['changed_since_last'])} -> {fi['changed_since_last'][:20]}")
    lines.append("")
    syn = report["sections"]["syntax"]
    lines.append(f"SYNTAX: OK={syn['ok']} FAIL={syn['fail']}")
    if syn["fails"]:
        lines.append(" - First failures:")
        for x in syn["fails"][:10]:
            lines.append(f"   * {x['file']}: {x['error']}")
    lines.append("")
    imp = report["sections"]["imports"]
    lines.append(f"IMPORTS: OK={len(imp['ok'])} FAIL={len(imp['fail'])}")
    if imp["fail"]:
        lines.append(" - First failures:")
        for kk,vv in list(imp["fail"].items())[:10]:
            lines.append(f"   * {kk}: {vv}")
    lines.append("")
    dd = report["sections"]["diagram_doctor"]
    lines.append(f"DIAGRAMS: OK={dd['ok']} FALLBACK={dd['fallback']} ERROR={dd['error']}")
    if dd["top_errors"]:
        lines.append(" - Top diagram errors:")
        for x in dd["top_errors"][:10]:
            lines.append(f"   * {x['key']}: {x['error']}")
    lines.append("")
    if "export_smoke" in report["sections"]:
        lines.append("EXPORT SMOKE:")
        lines.append(f" - {report['sections']['export_smoke']}")
        lines.append("")
    if "ingest_smoke" in report["sections"]:
        lines.append("INGEST SMOKE:")
        lines.append(f" - {report['sections']['ingest_smoke']}")
        lines.append("")

    txt_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("OK: wrote reports to", rep_dir)
    print((txt_path).read_text(encoding="utf-8"))

if __name__ == "__main__":
    main()
