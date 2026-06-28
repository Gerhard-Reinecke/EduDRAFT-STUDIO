# ======================================================================================
# diagram_library.py
# ======================================================================================
# Module: Diagram Resolution, Deterministic Rendering & Visual Fallback Engine
#
# System: EduDraft Studio (Marike App)
# Version: 3.0
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module is the structured visual rendering engine of EduDraft Studio.
#
# It converts diagram intent into clean, printable, classroom-ready visual outputs
# through a deterministic-first architecture supported by a controlled image-generation
# fallback pathway.
#
# This is not a generic drawing library. It is a governed diagram generation pipeline
# built around a Master Archetype Registry, stable request/response contracts, and
# explicit cost-control rules.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Intent Normalisation
#    - Converts upstream diagram intent into a stable request structure
#    - Normalises prompt, subject, level, archetype hints, and parameters
#    - Provides universal request builders for app and LLM integration
#
# 2. Archetype Resolution
#    - Resolves diagram intent to a single archetype_id
#    - Uses registry metadata, synonyms, subject matching, and level filtering
#    - Ensures that all supported visuals flow through a consistent classification model
#
# 3. Deterministic Rendering Layer
#    - Uses coded renderers (primarily matplotlib-based) for known diagram types
#    - Produces consistent, explainable, black-and-white classroom-safe visuals
#    - Keeps deterministic diagrams credit-free
#
# 4. Fallback Visual Generation
#    - Activates when:
#        • no archetype match is found
#        • renderer is not implemented
#        • archetype is marked as non-deterministic
#    - Uses image generation under explicit credit-control rules
#    - Never provides silent free generation when charging logic fails
#
# 5. Credit-Controlled Image Generation
#    - Enforces spend-before-generate behaviour for paid fallback visuals
#    - Blocks free image generation when spend hooks are unavailable
#    - Returns valid failure placeholder PNGs when image generation cannot proceed
#
# 6. Stable Output Contract
#    - Returns structured diagram responses in a predictable format:
#        {
#          "status": "ok" | "fallback" | "error",
#          "archetype_id": str | None,
#          "title": str | None,
#          "cost_credits": float,
#          "mime": "image/png" | None,
#          "bytes": b"...",
#          "debug": {...}
#        }
#
# --------------------------------------------------------------------------------------
# MASTER ARCHETYPE REGISTRY
# --------------------------------------------------------------------------------------
# The ARCHETYPES dictionary is the single source of truth for supported diagram types.
#
# Each archetype defines:
#   - title
#   - subjects
#   - supported levels
#   - deterministic / fallback behaviour
#   - renderer function
#   - required and optional parameters
#   - defaults by level
#   - synonym mappings
#
# Every diagram request should resolve to exactly one archetype wherever possible.
#
# --------------------------------------------------------------------------------------
# KEY SUBSYSTEMS
# --------------------------------------------------------------------------------------
# - Archetype registry
# - Intent-to-request bridge helpers
# - Deterministic renderer functions
# - Image generation fallback wrapper
# - Credit spend hook integration
# - Failure placeholder image generation
# - App-friendly public API for diagram generation
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - matplotlib → deterministic visual rendering
# - OpenAI image generation → fallback visual generation
# - Standard library utilities for parsing, temp files, IO, math, and typing
# - App-level credit spend injection via set_credit_spend_fn(...)
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Deterministic first:
#     Prefer coded, explainable visual rendering over AI generation
#
# - Cost control:
#     No free fallback image generation; credits enforced before execution
#
# - Graceful failure:
#     Always return valid PNG bytes, even when generation fails
#
# - Classroom safety:
#     Outputs are clean, printable, and education-appropriate
#
# - Subject agnostic:
#     Supports multiple domains including maths, physics, chemistry,
#     biology, geography, and related educational contexts
#
# - Extensible architecture:
#     New archetypes can be added without destabilising existing logic
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module is the single visual generation engine for the platform.
#
# In practical terms:
#   - llm.py defines where visuals are needed via placeholders
#   - exports.py routes placeholders into rendering flows
#   - diagram_library.py decides how the visual is actually produced
#
# It therefore directly affects:
#   - rendering quality
#   - cost control
#   - diagram consistency
#   - export stability
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This module is operationally critical across document, PPT, and export workflows
# - Any change here affects both deterministic diagrams and paid visual fallbacks
# - Credit enforcement must remain strict to preserve commercial control
# - Archetype registry design is foundational to long-term scalability
#
# ======================================================================================

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple
import io
import math
import re
import tempfile

import matplotlib
matplotlib.use("Agg")  # headless, HF safe
import matplotlib.pyplot as plt
from fractions import Fraction


# =============================
# IMAGE GEN + CREDITS (injected by app)
# =============================

import os
import base64
from typing import List, Optional, Dict, Any, Tuple

# app.py will inject a spend function here:
#   spend_fn(cost: float, reason: str, meta: dict) -> tuple[bool, str, float]
CREDIT_SPEND_FN = None

# default cost per generated image (paid fallback)
CREDIT_COST_PER_IMAGE = float(os.environ.get("IMAGE_CREDITS_PER_IMAGE", "1"))

def set_credit_spend_fn(fn):
    """
    app.py should call this once after it defines the user's spend function.
    """
    global CREDIT_SPEND_FN
    CREDIT_SPEND_FN = fn


def _image_gen_failed_png_bytes(where: str, err_text: str, prompt: str = "") -> bytes:
    """
    Always returns valid PNG BYTES (never empty).
    Used when image generation fails so exports NEVER crash.
    """
    fig, ax = plt.subplots(figsize=(8.27, 11.69), dpi=220)
    ax.axis("off")

    title = f"IMAGE (GEN FAILED) — {where or 'Unspecified'}"
    ax.text(0.0, 1.02, title, fontsize=16, weight="bold", va="bottom", transform=ax.transAxes)

    msg = (err_text or "").strip()
    pr = (prompt or "").strip()

    if len(msg) > 500:
        msg = msg[:500] + "..."
    if len(pr) > 800:
        pr = pr[:800] + "..."

    ax.text(0.0, 0.93, "Reason:", fontsize=12, weight="bold", va="top", transform=ax.transAxes)
    ax.text(0.0, 0.90, msg or "(unknown error)", fontsize=11, va="top", transform=ax.transAxes)

    if pr:
        ax.text(0.0, 0.78, "Prompt (truncated):", fontsize=12, weight="bold", va="top", transform=ax.transAxes)
        ax.text(0.0, 0.75, pr, fontsize=10, va="top", transform=ax.transAxes)

    # border
    ax.add_patch(plt.Rectangle((0, 0), 1, 1, fill=False, linewidth=1.2, transform=ax.transAxes))

    return _fig_to_png_bytes(fig)


def _should_allow_image_gen(user_ctx: Optional[Dict[str, Any]]) -> bool:
    """
    Central switch: disable image-gen globally if needed,
    or per-user (e.g., no credits / not logged in).
    """
    # If your app has a global toggle, keep it here:
    # (If you don't have it, leave as True)
    if os.environ.get("ENABLE_IMAGE_GEN", "1").strip() in {"0", "false", "False"}:
        return False
    return True


def render_image_gen_png_bytes(
    where: str,
    prompt: str,
    notes: str = "",
    user_ctx: Optional[Dict[str, Any]] = None,
    *,
    reason: str = "image_gen",
) -> Tuple[bytes, Dict[str, Any]]:
    """
    Generates a worksheet-style image via OpenAI image generation.
    Returns (png_bytes, meta).
    Raises RuntimeError when credits are insufficient so callers can block generation.
    IMPORTANT RULE:
      - Credits are spent BEFORE image generation.
      - If credit spend fails -> do NOT call OpenAI or allow free image.
    """
    user_ctx = user_ctx or {}
    title = (where or "").strip()
    user_prompt = (prompt or "").strip()
    user_notes = (notes or "").strip()

    final_prompt = (
        "Create a clear BLACK-AND-WHITE worksheet diagram suitable for printing.\n"
        "Style rules:\n"
        "- clean vector/line-art look\n"
        "- no shading, no gradients, no color\n"
        "- minimal background (white)\n"
        "- NO text labels unless explicitly requested\n"
        "- high clarity, classroom-safe\n\n"
        f"Audience level: {user_ctx.get('level','')}\n"
        f"Complexity tier: {user_ctx.get('tier','')}\n"
        "Match the diagram detail to this level.\n\n"
        f"Diagram topic/context: {title}\n\n"
        f"User request:\n{user_prompt}\n\n"
    )
    if user_notes:
        final_prompt += f"Notes:\n{user_notes}\n\n"

    # If image-gen not allowed, return a "disabled" placeholder (still valid image bytes)
    if not _should_allow_image_gen(user_ctx):
        return (
            _image_gen_failed_png_bytes(title or "Image", "Image generation is disabled.", user_prompt),
            {"charged": False, "cost": 0.0, "reason": "disabled"}
        )

    cost = float(CREDIT_COST_PER_IMAGE)

    if CREDIT_SPEND_FN is None:
        # If no spend fn is wired, do NOT give free images.
        return (
            _image_gen_failed_png_bytes(title or "Image", "Credit spend function is not wired (CREDIT_SPEND_FN is None).", user_prompt),
            {"charged": False, "cost": cost, "reason": "spend_fn_missing"}
        )

    try:
        ok, msg, _balance = CREDIT_SPEND_FN(cost, reason, {
            "where": title,
            "model": os.environ.get("OPENAI_IMAGE_MODEL", ""),
            "size": os.environ.get("OPENAI_IMAGE_SIZE", ""),
        })
    except Exception as e:
        return (
            _image_gen_failed_png_bytes(title or "Image", f"Credit spend failed: {type(e).__name__}: {e}", user_prompt),
            {"charged": False, "cost": cost, "reason": "spend_failed"}
        )

    if not ok:
        if "insufficient" in (msg or "").lower():
            raise RuntimeError("Not enough credits to generate images/diagrams.")
        return (
            _image_gen_failed_png_bytes(title or "Image", msg or "Credit spend failed.", user_prompt),
            {"charged": False, "cost": cost, "reason": "spend_failed"}
        )

    # Local import to avoid hard crash if OpenAI SDK isn't available in some environments
    try:
        from openai import OpenAI
    except Exception as e:
        return (
            _image_gen_failed_png_bytes(title or "Image", f"OpenAI SDK import failed: {type(e).__name__}: {e}", user_prompt),
            {"charged": False, "cost": 0.0, "reason": "sdk_missing"}
        )

    try:
        client = OpenAI()

        img = client.images.generate(
            model=os.environ.get("OPENAI_IMAGE_MODEL", "gpt-image-1"),
            prompt=final_prompt,
            n=1,
            size=os.environ.get("OPENAI_IMAGE_SIZE", "1024x1024"),
        )

        image_bytes = base64.b64decode(img.data[0].b64_json)

        return image_bytes, {"charged": True, "cost": cost, "reason": "ok"}

    except Exception as e:
        print("IMAGE_GEN_ERROR:", repr(e))
        return (
            _image_gen_failed_png_bytes(title or "Image", f"{type(e).__name__}: {e}", user_prompt),
            {"charged": False, "cost": 0.0, "reason": "gen_error"}
        )


# ADD THIS SECTION to diagram_library.py (Step 8)
# ----------------------------------------------
# Goal: Provide ONE universal wrapper that converts your LLM's diagram intent
# into the exact request dict shape this library expects.
#
# This does NOT try to magically parse everything (that belongs upstream).
# Instead it:
#   1) Accepts a structured "diagram_intent" object from your LLM pipeline
#   2) Normalizes fields
#   3) Produces a request dict
#
# You can plug this in immediately so your app always talks to the library
# in the same way.

def build_request_from_intent(diagram_intent: Dict[str, Any]) -> Dict[str, Any]:
    """
    Universal bridge: LLM intent -> diagram_library request.
    Expected diagram_intent (examples):
      {
        "prompt": "Create a histogram of test results",
        "subject": "math",
        "level": "S",
        "archetype_id": "histogram",
        "params": {"bins": [0,10,20,30], "frequencies": [3,5,2], "title": "Test results"}
      }
    Or minimal:
      {
        "prompt": "Labelled human heart",
        "subject": "biology",
        "level": "J"
      }
    Output request (guaranteed keys):
      {
        "prompt": str,
        "subject": str,
        "level": str,
        "archetype_hint": str|"" ,
        "params": dict
      }
    """
    prompt = (diagram_intent.get("prompt") or diagram_intent.get("text") or "").strip()
    subject = (diagram_intent.get("subject") or "").strip().lower()
    level = (diagram_intent.get("level") or diagram_intent.get("year_level") or "J").strip().upper()

    # Accept both "archetype_id" or "archetype"
    archetype_hint = (diagram_intent.get("archetype_id") or diagram_intent.get("archetype") or "").strip()

    params = diagram_intent.get("params") or diagram_intent.get("arguments") or {}
    if not isinstance(params, dict):
        # keep it safe; library expects dict
        params = {}

    # If the LLM provided an "items" style payload (common), map it into params if not already present.
    # This is intentionally conservative and only applies for a few high-ROI archetypes.
    if archetype_hint == "bar_chart":
        # allow: {"items":[{"category":"A","value":3}, ...]}
        if "categories" not in params and "values" not in params:
            items = diagram_intent.get("items")
            if isinstance(items, list) and items and isinstance(items[0], dict):
                cats = [str(it.get("category", "")).strip() for it in items]
                vals = [it.get("value") for it in items]
                if all(cats) and all(isinstance(v, (int, float)) for v in vals):
                    params["categories"] = cats
                    params["values"] = vals

    if archetype_hint == "line_graph":
        # allow: {"points":[{"x":1,"y":2}, ...]}
        if "x" not in params and "y" not in params:
            pts = diagram_intent.get("points")
            if isinstance(pts, list) and pts and isinstance(pts[0], dict):
                xs = [p.get("x") for p in pts]
                ys = [p.get("y") for p in pts]
                if all(isinstance(v, (int, float)) for v in xs) and all(isinstance(v, (int, float)) for v in ys):
                    params["x"] = xs
                    params["y"] = ys

    if archetype_hint == "scatter_plot":
        if "x" not in params and "y" not in params:
            pts = diagram_intent.get("points")
            if isinstance(pts, list) and pts and isinstance(pts[0], dict):
                xs = [p.get("x") for p in pts]
                ys = [p.get("y") for p in pts]
                if all(isinstance(v, (int, float)) for v in xs) and all(isinstance(v, (int, float)) for v in ys):
                    params["x"] = xs
                    params["y"] = ys

    if archetype_hint == "coordinate_plane_points":
        # allow: {"points":[[1,2,"A"], [3,4,"B"]]} or list of dicts
        if "points" not in params:
            pts = diagram_intent.get("points")
            if isinstance(pts, list) and pts:
                params["points"] = pts

    if archetype_hint == "histogram":
        # allow: {"bins":[...], "freq":[...]} alias mapping
        if "frequencies" not in params and "freq" in params:
            params["frequencies"] = params.pop("freq")

    if archetype_hint == "box_and_whisker":
        # allow common aliases
        alias = {
            "minimum": "min",
            "lower_quartile": "q1",
            "upper_quartile": "q3",
        }
        for k_src, k_dst in alias.items():
            if k_dst not in params and k_src in params:
                params[k_dst] = params[k_src]

    return {
        "prompt": prompt,
        "subject": subject,
        "level": level,
        "archetype_hint": archetype_hint,
        "params": params,
    }


# OPTIONAL: One-liner convenience call used by the app:
def generate_from_intent(diagram_intent: Dict[str, Any], user_ctx: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    App-friendly helper: takes LLM intent, builds request, calls generate_diagram.
    """
    req = build_request_from_intent(diagram_intent)
    return generate_diagram(req, user_ctx=user_ctx)


# ------------------------------------------------
# This implements a deterministic 2-stage probability tree renderer.
# It covers most school use-cases (two sequential events).
#
# INPUT FORMAT (required):
#   stages = [
#     {
#       "name": "A",
#       "branches": [
#         {"label": "Heads", "p": "1/2"},
#         {"label": "Tails", "p": "1/2"}
#       ]
#     },
#     {
#       "name": "B",
#       "branches": [
#         {"label": "Red", "p": "1/3"},
#         {"label": "Blue", "p": "2/3"}
#       ]
#     }
#   ]
#
# Optional:
#   title
#   show_joint (bool) -> show joint probability on leaves if both probs are parseable
#   p_format ("fraction"|"decimal") default "fraction"
#
# NOTE: This renderer draws a clean tree with text labels. It avoids fancy layout libs.


def render_probability_tree(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    stages = params["stages"]
    plot_title = params.get("title") or title or "Probability tree"
    show_joint = bool(params.get("show_joint", False))
    p_format = (params.get("p_format") or "fraction").strip().lower()

    if not isinstance(stages, list) or len(stages) != 2:
        raise ValueError("This renderer supports exactly 2 stages: stages must be a list of length 2.")

    s1 = stages[0]
    s2 = stages[1]
    if not isinstance(s1, dict) or not isinstance(s2, dict):
        raise ValueError("Each stage must be a dict with keys: name, branches.")
    b1 = s1.get("branches")
    b2 = s2.get("branches")
    if not (isinstance(b1, list) and b1) or not (isinstance(b2, list) and b2):
        raise ValueError("Each stage must have a non-empty 'branches' list.")

    # Parse probs if possible
    def parse_p(x) -> Optional[Fraction]:
        if x is None:
            return None
        if isinstance(x, (int, float)):
            # treat float as exact fraction where possible
            try:
                return Fraction(x).limit_denominator(1000000)
            except Exception:
                return None
        if isinstance(x, str):
            s = x.strip()
            # allow "1/2" or "0.25"
            try:
                if "/" in s:
                    return Fraction(s)
                return Fraction(float(s)).limit_denominator(1000000)
            except Exception:
                return None
        return None

    def fmt_p(fr: Optional[Fraction], original: Any) -> str:
        if fr is None:
            return str(original)
        if p_format == "decimal":
            return f"{float(fr):.3g}"
        # fraction default
        return f"{fr.numerator}/{fr.denominator}" if fr.denominator != 1 else str(fr.numerator)

    # Build labels
    stage1_labels = []
    for br in b1:
        if not isinstance(br, dict):
            raise ValueError("Each branch must be a dict like {'label':..., 'p':...}.")
        stage1_labels.append((str(br.get("label", "")).strip(), br.get("p", "")))

    stage2_labels = []
    for br in b2:
        if not isinstance(br, dict):
            raise ValueError("Each branch must be a dict like {'label':..., 'p':...}.")
        stage2_labels.append((str(br.get("label", "")).strip(), br.get("p", "")))

    if any(not lab for lab, _ in stage1_labels) or any(not lab for lab, _ in stage2_labels):
        raise ValueError("All branches must have non-empty 'label'.")

    # Layout coordinates
    # Root at x=0.05, stage1 split at x=0.35, stage2 leaves at x=0.75
    # y positions evenly spaced for leaves; stage1 nodes centered above their leaves.
    n1 = len(stage1_labels)
    n2 = len(stage2_labels)
    leaves = n1 * n2

    # y positions for each leaf
    y_top = 0.9
    y_bot = 0.1
    if leaves == 1:
        y_leaf = [0.5]
    else:
        step = (y_top - y_bot) / (leaves - 1)
        y_leaf = [y_top - i * step for i in range(leaves)]

    # Map each stage1 branch to its leaf indices
    leaf_indices_by_b1 = {}
    idx = 0
    for i in range(n1):
        leaf_indices_by_b1[i] = list(range(idx, idx + n2))
        idx += n2

    x_root = 0.05
    x_s1 = 0.35
    x_leaf = 0.75

    # stage1 y positions = average of its leaves
    y_s1 = []
    for i in range(n1):
        inds = leaf_indices_by_b1[i]
        y_s1.append(sum(y_leaf[j] for j in inds) / len(inds))

    fig, ax = plt.subplots(figsize=(8.6, max(3.8, 0.45 * leaves)))
    ax.set_title(plot_title)
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    # Draw root
    ax.plot([x_root], [0.5], marker="o")
    ax.text(x_root, 0.55, s1.get("name", "Stage 1"), ha="left", fontsize=10)

    # Draw stage1 nodes and branches from root
    for i, (lab1, p1) in enumerate(stage1_labels):
        ax.plot([x_s1], [y_s1[i]], marker="o")
        ax.plot([x_root, x_s1], [0.5, y_s1[i]], linewidth=1.8)

        fr1 = parse_p(p1)
        p1_text = fmt_p(fr1, p1)

        ax.text((x_root + x_s1) / 2, (0.5 + y_s1[i]) / 2 + 0.03, f"{lab1}", ha="center", fontsize=10)
        if p1_text:
            ax.text((x_root + x_s1) / 2, (0.5 + y_s1[i]) / 2 - 0.03, f"p={p1_text}", ha="center", fontsize=9)

    # Draw stage2 leaves from each stage1 node
    leaf_idx = 0
    for i, (lab1, p1) in enumerate(stage1_labels):
        fr1 = parse_p(p1)

        for j, (lab2, p2) in enumerate(stage2_labels):
            y = y_leaf[leaf_idx]
            ax.plot([x_leaf], [y], marker="o")
            ax.plot([x_s1, x_leaf], [y_s1[i], y], linewidth=1.6)

            fr2 = parse_p(p2)
            p2_text = fmt_p(fr2, p2)

            midx = (x_s1 + x_leaf) / 2
            midy = (y_s1[i] + y) / 2

            ax.text(midx, midy + 0.03, f"{lab2}", ha="center", fontsize=10)
            if p2_text:
                ax.text(midx, midy - 0.03, f"p={p2_text}", ha="center", fontsize=9)

            # Optional joint probability on leaf
            joint_text = ""
            if show_joint and fr1 is not None and fr2 is not None:
                frj = fr1 * fr2
                joint_text = fmt_p(frj, frj)
            if joint_text:
                ax.text(x_leaf + 0.02, y, f"joint={joint_text}", ha="left", va="center", fontsize=9)

            # Leaf label (combined)
            ax.text(x_leaf + 0.02, y + 0.05, f"{lab1} → {lab2}", ha="left", fontsize=9)

            leaf_idx += 1

    # Stage 2 label
    ax.text(x_s1 + 0.02, 0.95, s2.get("name", "Stage 2"), ha="left", fontsize=10)

    return _fig_to_png_bytes(fig)


# -----------------------------------------------------------------------------
# 1) Master Inventory Registry (SOURCE OF TRUTH)
# -----------------------------------------------------------------------------

ARCHETYPES: Dict[str, Dict[str, Any]] = {
    # =====================================================================
    # A) MATHS CORE (implemented in this file: bar_chart, line_graph,
    #                scatter_plot, pie_chart, number_line)
    # =====================================================================

    "number_line": {
        "title": "Number line",
        "subjects": ["math"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_number_line",
        "required_params": ["min", "max"],
        "optional_params": ["ticks", "mark_points", "title"],
        "defaults_by_level": {"P": {"ticks": 6, "mark_points": []}, "J": {"ticks": 9, "mark_points": []}},
        "synonyms": ["number line", "show on a number line", "plot on number line", "fractions on a number line"],
    },

    "bar_chart": {
        "title": "Bar chart",
        "subjects": ["math"],
        "levels": ["P", "J", "S"],
        "deterministic": True,
        "renderer": "render_bar_chart",
        "required_params": ["categories", "values"],
        "optional_params": ["title", "x_label", "y_label", "units", "show_grid", "rotate_x", "legend"],
        "defaults_by_level": {
            "P": {"show_grid": False, "rotate_x": 0, "legend": False},
            "J": {"show_grid": True,  "rotate_x": 0, "legend": False},
            "S": {"show_grid": True,  "rotate_x": 20, "legend": False},
        },
        "synonyms": ["bar chart", "bar graph", "column chart", "column graph", "compare categories"],
    },

    "slope_triangle": {
        "title": "Slope triangle",
        "subjects": ["math"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_slope_triangle",
        "synonyms": ["slope triangle", "gradient triangle", "rise run triangle"],
        "required_params": ["p1", "p2"],
        "optional_params": ["title", "show_grid", "label_rise", "label_run", "x_min", "x_max", "y_min", "y_max"],
    },

    "gradient_rise_run": {
        "title": "Gradient rise/run",
        "subjects": ["math"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_gradient_rise_run",
        "synonyms": ["rise run", "gradient rise run", "find gradient", "slope rise run"],
        "required_params": ["p1", "p2"],
        "optional_params": ["title", "show_grid", "show_values", "x_min", "x_max", "y_min", "y_max"],
    },

    "line_graph": {
        "title": "Line graph",
        "subjects": ["math", "science"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_line_graph",
        "required_params": ["x", "y"],
        "optional_params": ["title", "x_label", "y_label", "units_x", "units_y", "show_grid", "marker"],
        "defaults_by_level": {"J": {"show_grid": True, "marker": True}, "S": {"show_grid": True, "marker": False}},
        "synonyms": ["line graph", "line chart", "plot a line", "time series", "trend line"],
    },

    "scatter_plot": {
        "title": "Scatter plot",
        "subjects": ["math", "science"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_scatter_plot",
        "required_params": ["x", "y"],
        "optional_params": ["title", "x_label", "y_label", "units_x", "units_y", "show_grid", "best_fit_line"],
        "defaults_by_level": {"J": {"show_grid": True, "best_fit_line": False}, "S": {"show_grid": True, "best_fit_line": True}},
        "synonyms": ["scatter plot", "scatter graph", "correlation plot", "points on a graph"],
    },

        "scatter_plot_blank": {
        "title": "Scatter plot (blank grid template)",
        "subjects": ["math", "science"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_scatter_plot_blank",
        "required_params": [],
        "optional_params": ["title", "x_label", "y_label", "x_min", "x_max", "y_min", "y_max", "x_step", "y_step", "show_grid"],
        "defaults_by_level": {
            "J": {"x_min": 0, "x_max": 10, "y_min": 0, "y_max": 10, "x_step": 1, "y_step": 1, "show_grid": True},
            "S": {"x_min": 0, "x_max": 20, "y_min": 0, "y_max": 20, "x_step": 2, "y_step": 2, "show_grid": True},
            "U": {"x_min": -50, "x_max": 50, "y_min": -50, "y_max": 50, "x_step": 10, "y_step": 10, "show_grid": True},
        },
        "synonyms": [
            "blank scatter plot",
            "scatter plot template",
            "scatter graph template",
            "use the grid below",
            "plot points on the grid",
            "provide space for students"
        ],
    },

    "pie_chart": {
        "title": "Pie chart",
        "subjects": ["math"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_pie_chart",
        "required_params": ["labels", "values"],
        "optional_params": ["title", "show_percent"],
        "defaults_by_level": {"P": {"show_percent": True}, "J": {"show_percent": True}},
        "synonyms": ["pie chart", "pie graph", "sector chart"],
    },

    # --- Maths archetypes (registered now; renderer to add later) ---
    "histogram": {
        "title": "Histogram",
        "subjects": ["math"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_histogram",  # not yet implemented -> clean fallback
        "required_params": ["bins", "frequencies"],
        "optional_params": ["title", "x_label", "y_label", "show_grid"],
        "defaults_by_level": {"J": {"show_grid": True}, "S": {"show_grid": True}},
        "synonyms": ["histogram", "grouped frequency histogram", "binned data histogram"],
    },

        "histogram_blank_axes": {
        "title": "Histogram (blank axes template)",
        "subjects": ["math"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_histogram_blank_axes",
        "required_params": [],
        "optional_params": ["bins", "frequencies", "title", "x_label", "y_label", "show_grid", "y_max"],
        "defaults_by_level": {"J": {"show_grid": True}, "S": {"show_grid": True}},
        "synonyms": [
            "blank histogram axes",
            "blank set of axes",
            "provide space for students to draw",
            "draw a histogram on the axes",
            "use the axes below",
            "construct a histogram"
        ],
    },

    "box_and_whisker": {
        "title": "Box and whisker plot",
        "subjects": ["math"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_box_and_whisker",  # not yet implemented
        "required_params": ["min", "q1", "median", "q3", "max"],
        "optional_params": ["title", "show_outliers", "outliers"],
        "defaults_by_level": {"J": {"show_outliers": False, "outliers": []}, "S": {"show_outliers": True, "outliers": []}},
        "synonyms": ["box plot", "box and whisker", "quartile plot"],
    },

        "box_and_whisker_blank": {
        "title": "Box and whisker plot (blank template)",
        "subjects": ["math"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_box_and_whisker_blank",
        "required_params": [],
        "optional_params": ["title", "x_min", "x_max", "x_step", "show_grid", "show_key"],
        "defaults_by_level": {
            "J": {"x_min": 0, "x_max": 40, "x_step": 5, "show_grid": True, "show_key": True},
            "S": {"x_min": 0, "x_max": 100, "x_step": 10, "show_grid": True, "show_key": True},
            "U": {"x_min": 0, "x_max": 200, "x_step": 20, "show_grid": True, "show_key": True},
        },
        "synonyms": [
            "blank box plot",
            "box and whisker template",
            "quartile plot template",
            "provide space for students",
            "use the axis below"
        ],
    },

    "stem_and_leaf": {
        "title": "Stem-and-leaf plot",
        "subjects": ["math"],
        "levels": ["J"],
        "deterministic": True,
        "renderer": "render_stem_and_leaf",  # not yet implemented
        "required_params": ["data"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}},
        "synonyms": ["stem and leaf", "stem-and-leaf", "stem leaf plot"],
    },

    "stem_and_leaf_blank": {
        "title": "Stem-and-leaf plot (blank template)",
        "subjects": ["math"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_stem_and_leaf_blank",
        "required_params": [],
        "optional_params": ["title", "stems", "min", "max", "stem_unit", "show_key", "key_example", "rows"],
        "defaults_by_level": {
            "J": {"stem_unit": 10, "rows": 12, "show_key": True},
            "S": {"stem_unit": 10, "rows": 18, "show_key": True},
            "U": {"stem_unit": 10, "rows": 22, "show_key": True},
        },
        "synonyms": [
            "blank stem and leaf",
            "stem-and-leaf template",
            "provide space for stem and leaf",
            "stem and leaf table",
            "construct a stem-and-leaf plot"
        ],
    },

    "frequency_table": {
        "title": "Frequency table",
        "subjects": ["math"],
        "levels": ["P", "J", "S"],
        "deterministic": True,
        "renderer": "render_frequency_table",  # not yet implemented
        "required_params": ["rows"],  # list of {"label":..., "freq":...} or similar
        "optional_params": ["title", "show_tally"],
        "defaults_by_level": {"P": {"show_tally": True}, "J": {"show_tally": False}, "S": {"show_tally": False}},
        "synonyms": ["frequency table", "tally table", "count table"],
    },

        "frequency_table_blank": {
        "title": "Frequency table (blank template)",
        "subjects": ["math"],
        "levels": ["P", "J", "S", "U"],
        "deterministic": True,
        "renderer": "render_frequency_table_blank",
        "required_params": [],
        "optional_params": ["title", "show_tally", "show_total", "row_count", "label_header", "freq_header", "tally_header"],
        "defaults_by_level": {
            "P": {"show_tally": True,  "show_total": False, "row_count": 6,  "label_header": "Category", "freq_header": "Frequency", "tally_header": "Tally"},
            "J": {"show_tally": False, "show_total": False, "row_count": 8,  "label_header": "Category", "freq_header": "Frequency", "tally_header": "Tally"},
            "S": {"show_tally": False, "show_total": False, "row_count": 10, "label_header": "Category", "freq_header": "Frequency", "tally_header": "Tally"},
            "U": {"show_tally": False, "show_total": False, "row_count": 12, "label_header": "Category", "freq_header": "Frequency", "tally_header": "Tally"},
        },
        "synonyms": [
            "blank frequency table",
            "frequency table template",
            "tally table template",
            "provide a frequency table",
            "provide space for a frequency table",
            "fill in the frequency table"
        ],
    },

    "two_way_table": {
        "title": "Two-way table",
        "subjects": ["math"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_two_way_table",  # not yet implemented
        "required_params": ["row_labels", "col_labels", "values_matrix"],
        "optional_params": ["title", "show_totals"],
        "defaults_by_level": {"J": {"show_totals": True}, "S": {"show_totals": True}},
        "synonyms": ["two-way table", "contingency table", "cross tabulation"],
    },

        "two_way_table_blank": {
        "title": "Two-way table (blank template)",
        "subjects": ["math"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_two_way_table_blank",
        "required_params": [],
        "optional_params": ["title", "row_labels", "col_labels", "show_totals", "corner_label", "row_total_label", "col_total_label", "row_count", "col_count"],
        "defaults_by_level": {
            "J": {"show_totals": True, "row_count": 2, "col_count": 2, "corner_label": "", "row_total_label": "Total", "col_total_label": "Total"},
            "S": {"show_totals": True, "row_count": 3, "col_count": 3, "corner_label": "", "row_total_label": "Total", "col_total_label": "Total"},
            "U": {"show_totals": True, "row_count": 4, "col_count": 4, "corner_label": "", "row_total_label": "Total", "col_total_label": "Total"},
        },
        "synonyms": [
            "blank two-way table",
            "two-way table template",
            "contingency table template",
            "provide a two-way table",
            "fill in the two-way table",
            "cross tabulation template"
        ],
    },

    "probability_tree": {
        "title": "Probability tree",
        "subjects": ["math"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_probability_tree",  # not yet implemented
        "required_params": ["stages"],  # structured tree data
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["probability tree", "tree diagram probability", "branching probabilities"],
    },

    "venn_2": {
        "title": "Venn diagram (2 sets)",
        "subjects": ["math"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_venn_2",  # not yet implemented
        "required_params": ["a_only", "b_only", "both"],
        "optional_params": ["title", "label_a", "label_b"],
        "defaults_by_level": {"J": {"label_a": "A", "label_b": "B"}, "S": {"label_a": "A", "label_b": "B"}},
        "synonyms": ["venn diagram", "2-set venn", "two circle venn"],
    },

    "venn_3": {
        "title": "Venn diagram (3 sets)",
        "subjects": ["math"],
        "levels": ["S"],
        "deterministic": True,
        "renderer": "render_venn_3",  # not yet implemented
        "required_params": ["regions"],  # dict for 7 regions
        "optional_params": ["title", "labels"],
        "defaults_by_level": {"S": {"labels": ["A", "B", "C"]}},
        "synonyms": ["3-set venn", "three circle venn", "venn diagram 3 sets"],
    },

    "coordinate_plane_points": {
        "title": "Coordinate plane (points)",
        "subjects": ["math"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_coordinate_plane_points",  # not yet implemented
        "required_params": ["points"],  # list of {"x":..,"y":..,"label":..}
        "optional_params": ["title", "x_min", "x_max", "y_min", "y_max", "show_grid"],
        "defaults_by_level": {"P": {"show_grid": True}, "J": {"show_grid": True}},
        "synonyms": ["plot points", "coordinate plane", "cartesian plane points"],
    },

    "linear_function_plot": {
        "title": "Linear function plot",
        "subjects": ["math"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_linear_function_plot",  # not yet implemented
        "required_params": ["m", "b"],
        "optional_params": ["title", "x_min", "x_max", "show_grid"],
        "defaults_by_level": {"J": {"show_grid": True}, "S": {"show_grid": True}},
        "synonyms": ["y=mx+b", "linear function graph", "plot a linear function"],
    },

    "quadratic_plot": {
        "title": "Quadratic function plot",
        "subjects": ["math"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_quadratic_plot",  # not yet implemented
        "required_params": ["a", "b", "c"],
        "optional_params": ["title", "x_min", "x_max", "show_grid"],
        "defaults_by_level": {"S": {"show_grid": True}, "U": {"show_grid": True}},
        "synonyms": ["quadratic graph", "parabola", "ax^2+bx+c graph"],
    },

    "trig_graphs": {
        "title": "Trigonometric graphs",
        "subjects": ["math"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_trig_graphs",  # not yet implemented
        "required_params": ["functions"],  # e.g. ["sin","cos"]
        "optional_params": ["title", "x_min", "x_max", "show_grid"],
        "defaults_by_level": {"S": {"show_grid": True}, "U": {"show_grid": True}},
        "synonyms": ["sine graph", "cosine graph", "tangent graph", "trig graphs"],
    },

    "angle_diagram": {
        "title": "Angle diagram",
        "subjects": ["math"],
        "levels": ["P", "J", "S"],
        "deterministic": True,
        "renderer": "render_angle_diagram",  # not yet implemented
        "required_params": ["type"],  # e.g. "parallel_lines", "triangle", etc
        "optional_params": ["title", "labels"],
        "defaults_by_level": {"P": {}, "J": {}, "S": {}},
        "synonyms": ["angle diagram", "angles on parallel lines", "angle relationships"],
    },

    "triangle_diagram": {
        "title": "Triangle diagram",
        "subjects": ["math"],
        "levels": ["P", "J", "S"],
        "deterministic": True,
        "renderer": "render_triangle_diagram",  # not yet implemented
        "required_params": ["labels"],  # sides/angles labels
        "optional_params": ["title", "show_right_angle"],
        "defaults_by_level": {"P": {"show_right_angle": False}, "J": {"show_right_angle": False}, "S": {"show_right_angle": False}},
        "synonyms": ["triangle diagram", "labelled triangle", "triangle with sides and angles"],
    },

    "circle_geometry": {
        "title": "Circle geometry diagram",
        "subjects": ["math"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_circle_geometry",  # not yet implemented
        "required_params": ["case"],  # chord/tangent/central angle etc
        "optional_params": ["title", "labels"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["circle geometry", "tangent chord theorem", "circle theorem diagram"],
    },

    "area_perimeter_shapes": {
        "title": "Area/perimeter shapes diagram",
        "subjects": ["math"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_area_perimeter_shapes",
        "required_params": [],
        "optional_params": ["title", "prompt", "description", "notes", "shapes"],
        "defaults_by_level": {"P": {}, "J": {}},
        "synonyms": [
            "area and perimeter",
            "composite shapes",
            "shape with dimensions",
            "parallelogram",
            "trapezium",
            "trapezoid",
            "circle diagram",
            "bowl option",
            "kite",
            "pythagoras",
            "safety rail",
            "sloped bank",
        ],
    },

    "net_3d_shapes": {
        "title": "3D nets",
        "subjects": ["math"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_net_3d_shapes",  # not yet implemented
        "required_params": ["solid"],  # cube/prism/pyramid
        "optional_params": ["title"],
        "defaults_by_level": {"P": {}, "J": {}},
        "synonyms": ["net of a cube", "3d net", "nets of solids"],
    },

    "solid_3d": {
        "title": "3D solid with dimensions",
        "subjects": ["math"],
        "levels": ["P", "J", "S"],
        "deterministic": True,
        "renderer": "render_solid_3d",  # not yet implemented
        "required_params": ["solid", "dimensions"],
        "optional_params": ["title"],
        "defaults_by_level": {"P": {}, "J": {}, "S": {}},
        "synonyms": ["3d solid", "rectangular prism", "cylinder with dimensions"],
    },

    # =====================================================================
    # B) PHYSICS / ENGINEERING (implemented: free_body_diagram, simple/series/parallel circuits)
    # =====================================================================

    "free_body_diagram": {
        "title": "Free body diagram",
        "subjects": ["physics", "engineering"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_free_body_diagram",
        "required_params": ["forces"],
        "optional_params": ["title", "show_axes"],
        "defaults_by_level": {"J": {"show_axes": True}, "S": {"show_axes": True}, "U": {"show_axes": True}},
        "synonyms": ["free body diagram", "fbd", "forces on object", "draw the forces"],
    },

    "motion_diagram": {
        "title": "Motion diagram (dot diagram / vectors)",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_motion_diagram",  # not yet implemented
        "required_params": ["positions"],
        "optional_params": ["title", "show_velocity_vectors"],
        "defaults_by_level": {"J": {"show_velocity_vectors": False}, "S": {"show_velocity_vectors": True}},
        "synonyms": ["motion diagram", "dot diagram motion", "displacement vectors diagram"],
    },

    "velocity_time_graph": {
        "title": "Velocity-time graph",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_velocity_time_graph",  # not yet implemented
        "required_params": ["t", "v"],
        "optional_params": ["title", "show_grid"],
        "defaults_by_level": {"J": {"show_grid": True}, "S": {"show_grid": True}},
        "synonyms": ["velocity time graph", "v-t graph", "speed time graph"],
    },

    "distance_time_graph": {
        "title": "Distance-time graph",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_distance_time_graph",  # not yet implemented
        "required_params": ["t", "s"],
        "optional_params": ["title", "show_grid"],
        "defaults_by_level": {"J": {"show_grid": True}, "S": {"show_grid": True}},
        "synonyms": ["distance time graph", "displacement time graph", "s-t graph"],
    },

    "energy_bar_chart": {
        "title": "Energy bar chart",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_energy_bar_chart",  # not yet implemented
        "required_params": ["stages"],  # list of stage dicts
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["energy bar chart", "energy transfer bars"],
    },

    "simple_circuit": {
        "title": "Simple circuit (cell + bulb + switch)",
        "subjects": ["science", "physics"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_simple_circuit",
        "required_params": [],
        "optional_params": ["title", "closed_switch"],
        "defaults_by_level": {"P": {"closed_switch": True}, "J": {"closed_switch": True}},
        "synonyms": ["simple circuit", "battery bulb switch", "cell bulb switch"],
    },

    "series_circuit": {
        "title": "Series circuit",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_series_circuit",
        "required_params": ["loads"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["series circuit", "components in series"],
    },

    "parallel_circuit": {
        "title": "Parallel circuit",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_parallel_circuit",
        "required_params": ["branches"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["parallel circuit", "components in parallel"],
    },

    "series_parallel_circuit": {
        "title": "Series-parallel circuit",
        "subjects": ["physics", "engineering"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_series_parallel_circuit",  # not yet implemented
        "required_params": ["network"],  # structured
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["series parallel circuit", "mixed circuit", "combination circuit"],
    },

    "circuit_symbols_sheet": {
        "title": "Circuit symbols reference sheet",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_circuit_symbols_sheet",  # not yet implemented
        "required_params": [],
        "optional_params": ["title", "symbols"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["circuit symbols", "electric circuit symbols sheet"],
    },

    "circuit_series_parallel": {
        "title": "Circuit (series/parallel)",
        "subjects": ["physics", "science", "engineering"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_circuit_series_parallel",
        "synonyms": ["series circuit", "parallel circuit", "circuit diagram", "battery lamp switch", "resistor circuit"],
        "required_params": ["mode", "components"],
        "optional_params": ["title", "show_switch", "switch_closed", "show_labels", "battery_label", "voltage_label", "annotate_current"],
    },

    "ray_diagram_reflection": {
        "title": "Ray diagram (reflection)",
        "subjects": ["physics"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_ray_diagram_reflection",  # not yet implemented
        "required_params": ["incident_angle_deg"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["ray diagram reflection", "mirror ray diagram", "reflection rays"],
    },

    "ray_diagram_refraction": {
        "title": "Ray diagram (refraction)",
        "subjects": ["physics"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_ray_diagram_refraction",  # not yet implemented
        "required_params": ["n1", "n2", "incident_angle_deg"],
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["ray diagram refraction", "snell's law diagram", "refraction rays"],
    },

    "lens_diagram": {
        "title": "Lens ray diagram",
        "subjects": ["physics"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_lens_diagram",  # not yet implemented
        "required_params": ["lens_type", "object_distance"],
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["lens diagram", "convex lens rays", "concave lens rays"],
    },

    "vector_addition": {
        "title": "Vector addition diagram",
        "subjects": ["math", "physics"],
        "levels": ["J", "S", "U"],
        "deterministic": True,
        "renderer": "render_vector_addition",  # not yet implemented
        "required_params": ["vectors"],  # list of {"angle_deg":..,"magnitude":..,"label":..}
        "optional_params": ["title", "method"],
        "defaults_by_level": {"J": {"method": "head_to_tail"}, "S": {"method": "head_to_tail"}, "U": {"method": "parallelogram"}},
        "synonyms": ["vector addition", "head to tail vectors", "parallelogram method"],
    },

    "inclined_plane_forces": {
        "title": "Inclined plane forces",
        "subjects": ["physics"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_inclined_plane_forces",  # not yet implemented
        "required_params": ["angle_deg", "mass"],
        "optional_params": ["title", "include_friction"],
        "defaults_by_level": {"S": {"include_friction": False}, "U": {"include_friction": True}},
        "synonyms": ["inclined plane forces", "block on slope forces", "ramp forces"],
    },

    "pulley_system": {
        "title": "Pulley system",
        "subjects": ["physics"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_pulley_system",  # not yet implemented
        "required_params": ["pulleys", "masses"],
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["pulley system", "block and tackle", "pulley diagram"],
    },

    # =====================================================================
    # C) CHEMISTRY (symbolic = deterministic; 3D molecules = fallback)
    # =====================================================================

    "periodic_table_excerpt": {
        "title": "Periodic table excerpt (highlighted groups/periods)",
        "subjects": ["chemistry"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_periodic_table_excerpt",  # not yet implemented
        "required_params": ["highlight"],  # e.g. {"group":17} or {"elements":["Na","Cl"]}
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["periodic table highlight", "highlight group", "highlight period", "periodic table excerpt"],
    },

    "atomic_structure_bohr": {
        "title": "Bohr model (shells and electrons)",
        "subjects": ["chemistry"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_atomic_structure_bohr",  # not yet implemented
        "required_params": ["symbol", "electrons_by_shell"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["bohr model", "shell diagram", "electron shells diagram"],
    },

    "electron_configuration_boxes": {
        "title": "Electron configuration (orbital box diagram)",
        "subjects": ["chemistry"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_electron_configuration_boxes",  # not yet implemented
        "required_params": ["configuration"],  # e.g. "1s2 2s2 2p6 ..."
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["orbital box diagram", "electron configuration boxes", "aufbau diagram"],
    },

    "lewis_dot": {
        "title": "Lewis dot structure",
        "subjects": ["chemistry"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_lewis_dot",  # not yet implemented
        "required_params": ["formula"],
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["lewis dot", "lewis structure", "dot and cross diagram"],
    },

    "reaction_energy_profile": {
        "title": "Reaction energy profile (endo/exo)",
        "subjects": ["chemistry"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_reaction_energy_profile",  # not yet implemented
        "required_params": ["type"],  # "endo" or "exo"
        "optional_params": ["title", "activation_energy", "delta_h"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["energy profile diagram", "reaction coordinate diagram", "endo exo energy graph"],
    },

    "titration_curve": {
        "title": "Titration curve",
        "subjects": ["chemistry"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_titration_curve",  # not yet implemented
        "required_params": ["type"],  # e.g. "strong_acid_strong_base"
        "optional_params": ["title", "equivalence_point"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["titration curve", "ph volume curve", "acid base titration graph"],
    },

    "ph_scale": {
        "title": "pH scale",
        "subjects": ["chemistry"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_ph_scale",  # not yet implemented
        "required_params": [],
        "optional_params": ["title", "markers"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["ph scale", "acidic neutral basic scale"],
    },

    "apparatus_setup_simple": {
        "title": "Simple apparatus setup (symbolic)",
        "subjects": ["chemistry"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_apparatus_setup_simple",  # not yet implemented
        "required_params": ["setup_type"],  # e.g. "filtration", "distillation_basic"
        "optional_params": ["title", "labels"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["lab apparatus diagram", "filtration setup", "distillation setup"],
    },

    "molecular_3d_ball_stick": {
        "title": "3D molecule (ball-and-stick)",
        "subjects": ["chemistry"],
        "levels": ["S", "U"],
        "deterministic": False,
        "fallback": "image_gen",
        "synonyms": ["ball and stick model", "3d molecule", "molecular model 3d"],
    },

    # =====================================================================
    # D) BIOLOGY (symbolic flow = deterministic; anatomy realism = fallback)
    # =====================================================================

    "food_web": {
        "title": "Food web",
        "subjects": ["biology"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_food_web",  # not yet implemented
        "required_params": ["nodes", "links"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["food web", "ecosystem food web", "trophic links diagram"],
    },

    "classification_tree": {
        "title": "Classification / taxonomy tree",
        "subjects": ["biology"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_classification_tree",  # not yet implemented
        "required_params": ["tree"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["taxonomy tree", "classification tree", "kingdom phylum class order family genus species"],
    },

    "cell_plant_simple": {
        "title": "Plant cell (simple labelled)",
        "subjects": ["biology"],
        "levels": ["J"],
        "deterministic": True,
        "renderer": "render_cell_plant_simple",  # not yet implemented (optional later)
        "required_params": ["labels"],  # allow custom label list
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}},
        "synonyms": ["plant cell diagram", "labelled plant cell"],
    },

    "cell_animal_simple": {
        "title": "Animal cell (simple labelled)",
        "subjects": ["biology"],
        "levels": ["J"],
        "deterministic": True,
        "renderer": "render_cell_animal_simple",  # not yet implemented
        "required_params": ["labels"],
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}},
        "synonyms": ["animal cell diagram", "labelled animal cell"],
    },

    "mitosis_stages": {
        "title": "Mitosis stages (symbolic)",
        "subjects": ["biology"],
        "levels": ["S"],
        "deterministic": True,
        "renderer": "render_mitosis_stages",  # not yet implemented
        "required_params": ["stages"],  # list of stage names
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}},
        "synonyms": ["mitosis stages", "prophase metaphase anaphase telophase diagram"],
    },

    "heart_labelled": {
        "title": "Labelled human heart",
        "subjects": ["biology"],
        "levels": ["J", "S"],
        "deterministic": False,
        "fallback": "image_gen",
        "synonyms": ["heart diagram", "human heart", "label the heart", "cardiac anatomy"],
    },

    "lungs_labelled": {
        "title": "Labelled lungs",
        "subjects": ["biology"],
        "levels": ["J", "S"],
        "deterministic": False,
        "fallback": "image_gen",
        "synonyms": ["lungs diagram", "respiratory system lungs", "label the lungs"],
    },

    "digestive_system": {
        "title": "Digestive system anatomy",
        "subjects": ["biology"],
        "levels": ["J", "S"],
        "deterministic": False,
        "fallback": "image_gen",
        "synonyms": ["digestive system diagram", "label the digestive system", "human digestion organs"],
    },

    # =====================================================================
    # E) GEOGRAPHY / EARTH SCIENCE
    # =====================================================================

    "compass_rose": {
        "title": "Compass rose",
        "subjects": ["geography"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_compass_rose",  # not yet implemented (you already had one earlier)
        "required_params": [],
        "optional_params": ["title"],
        "defaults_by_level": {"P": {}, "J": {}},
        "synonyms": ["compass rose", "compass directions", "north south east west"],
    },

    "map_key_legend": {
        "title": "Map key / legend",
        "subjects": ["geography"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_map_key_legend",  # not yet implemented
        "required_params": ["items"],  # list of {"symbol":"...","label":"..."}
        "optional_params": ["title"],
        "defaults_by_level": {"P": {}, "J": {}},
        "synonyms": ["map key", "map legend", "legend symbols"],
    },

    "scale_bar": {
        "title": "Scale bar",
        "subjects": ["geography"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_scale_bar",  # not yet implemented
        "required_params": ["length_km"],
        "optional_params": ["title", "segments"],
        "defaults_by_level": {"P": {"segments": 4}, "J": {"segments": 5}},
        "synonyms": ["scale bar", "map scale bar", "distance scale"],
    },

    "water_cycle": {
        "title": "Water cycle (symbolic)",
        "subjects": ["geography", "science"],
        "levels": ["P", "J"],
        "deterministic": True,
        "renderer": "render_water_cycle",  # not yet implemented
        "required_params": [],
        "optional_params": ["title", "labels"],
        "defaults_by_level": {"P": {}, "J": {}},
        "synonyms": ["water cycle", "evaporation condensation precipitation diagram"],
    },

    "rock_cycle": {
        "title": "Rock cycle (symbolic)",
        "subjects": ["geography", "science"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_rock_cycle",  # not yet implemented
        "required_params": [],
        "optional_params": ["title", "labels"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["rock cycle", "igneous sedimentary metamorphic cycle diagram"],
    },

    "plate_boundaries": {
        "title": "Plate boundaries (symbolic cross-sections)",
        "subjects": ["geography", "science"],
        "levels": ["J", "S"],
        "deterministic": True,
        "renderer": "render_plate_boundaries",  # not yet implemented
        "required_params": ["type"],  # "convergent" | "divergent" | "transform" | "subduction"
        "optional_params": ["title"],
        "defaults_by_level": {"J": {}, "S": {}},
        "synonyms": ["plate boundary diagram", "subduction zone diagram", "divergent boundary", "transform boundary"],
    },

    "contour_map_simple": {
        "title": "Simple contour map",
        "subjects": ["geography"],
        "levels": ["S", "U"],
        "deterministic": True,
        "renderer": "render_contour_map_simple",  # not yet implemented
        "required_params": ["contours"],  # list of polylines or simplified representation
        "optional_params": ["title"],
        "defaults_by_level": {"S": {}, "U": {}},
        "synonyms": ["contour map", "topographic contours", "contour lines map"],
    },

    "river_system_labelled": {
        "title": "Labelled river system",
        "subjects": ["geography"],
        "levels": ["P", "J"],
        "deterministic": False,
        "fallback": "image_gen",
        "synonyms": ["river diagram", "river system", "label a river", "meander oxbow delta"],
    },
}


# -----------------------------------------------------------------------------
# Area / perimeter geometry renderer
# -----------------------------------------------------------------------------

def render_area_perimeter_shapes(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic renderer for common school area/perimeter diagrams.

    Supports prompt-driven diagrams for:
    - parallelogram
    - trapezium / trapezoid
    - circles / bowl options
    - kite / composite kite joins
    - right-triangle / Pythagoras slope
    - simple side-by-side design overview

    This renderer is intentionally conservative. It creates clean black-and-white
    classroom-safe diagrams without using paid image generation.
    """
    plot_title = (params.get("title") or title or "Area/perimeter diagram").strip()
    prompt = str(params.get("prompt") or params.get("description") or params.get("notes") or "").lower()

    shapes = params.get("shapes")
    if isinstance(shapes, list) and shapes:
        prompt += " " + " ".join(str(s) for s in shapes).lower()

    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title(plot_title)

    def _label(x, y, text, size=10, weight="normal"):
        ax.text(x, y, str(text), ha="center", va="center", fontsize=size, weight=weight)

    def _line(x1, y1, x2, y2, label=None, label_offset=(0, 0)):
        ax.plot([x1, x2], [y1, y2], linewidth=1.6)
        if label:
            _label((x1 + x2) / 2 + label_offset[0], (y1 + y2) / 2 + label_offset[1], label)

    def _right_angle(x, y, size=0.25):
        ax.plot([x, x + size, x + size], [y + size, y + size, y], linewidth=1.0)

    def _draw_parallelogram(cx=0, cy=0, scale=1.0, title_text="Design A base"):
        pts = [
            (cx - 2.2 * scale, cy - 0.9 * scale),
            (cx + 1.8 * scale, cy - 0.9 * scale),
            (cx + 2.5 * scale, cy + 0.9 * scale),
            (cx - 1.5 * scale, cy + 0.9 * scale),
            (cx - 2.2 * scale, cy - 0.9 * scale),
        ]
        xs, ys = zip(*pts)
        ax.plot(xs, ys, linewidth=1.8)
        _label(cx, cy + 1.45 * scale, title_text, size=11, weight="bold")
        _line(cx - 2.2 * scale, cy - 0.9 * scale, cx + 1.8 * scale, cy - 0.9 * scale, "18 m", (0, -0.28 * scale))
        _line(cx - 2.2 * scale, cy - 0.9 * scale, cx - 1.5 * scale, cy + 0.9 * scale, "10 m", (-0.35 * scale, 0))
        _line(cx - 1.5 * scale, cy + 0.9 * scale, cx - 1.5 * scale, cy - 0.9 * scale, "7 m", (-0.28 * scale, 0))
        _right_angle(cx - 1.5 * scale, cy - 0.9 * scale, 0.22 * scale)

    def _draw_trapezium(cx=0, cy=0, scale=1.0, title_text="Design B base"):
        pts = [
            (cx - 2.4 * scale, cy - 0.9 * scale),
            (cx + 2.4 * scale, cy - 0.9 * scale),
            (cx + 1.1 * scale, cy + 0.9 * scale),
            (cx - 1.1 * scale, cy + 0.9 * scale),
            (cx - 2.4 * scale, cy - 0.9 * scale),
        ]
        xs, ys = zip(*pts)
        ax.plot(xs, ys, linewidth=1.8)
        _label(cx, cy + 1.45 * scale, title_text, size=11, weight="bold")
        _line(cx - 1.1 * scale, cy + 0.9 * scale, cx + 1.1 * scale, cy + 0.9 * scale, "10 m", (0, 0.25 * scale))
        _line(cx - 2.4 * scale, cy - 0.9 * scale, cx + 2.4 * scale, cy - 0.9 * scale, "22 m", (0, -0.28 * scale))
        _line(cx - 2.4 * scale, cy - 0.9 * scale, cx - 1.1 * scale, cy + 0.9 * scale, "8 m", (-0.35 * scale, 0))
        _line(cx + 2.4 * scale, cy - 0.9 * scale, cx + 1.1 * scale, cy + 0.9 * scale, "8 m", (0.35 * scale, 0))
        _line(cx, cy - 0.9 * scale, cx, cy + 0.9 * scale, "6 m", (0.25 * scale, 0))
        _right_angle(cx, cy - 0.9 * scale, 0.22 * scale)

    def _draw_circles():
        for cx, label, radius_text in [(-1.8, "Bowl Option 1", "r = 3 m"), (1.8, "Bowl Option 2", "r = 4 m")]:
            circle = plt.Circle((cx, 0), 0.95, fill=False, linewidth=1.8)
            ax.add_patch(circle)
            ax.plot([cx, cx + 0.95], [0, 0], linewidth=1.2)
            _label(cx, 1.35, label, size=11, weight="bold")
            _label(cx + 0.45, 0.18, radius_text, size=10)
            _label(cx, -1.35, "circle bowl", size=9)

    def _draw_kite(cx=0, cy=0, scale=1.0):
        pts = [
            (cx, cy + 1.5 * scale),
            (cx + 1.3 * scale, cy),
            (cx, cy - 1.5 * scale),
            (cx - 0.9 * scale, cy),
            (cx, cy + 1.5 * scale),
        ]
        xs, ys = zip(*pts)
        ax.plot(xs, ys, linewidth=1.8)
        ax.plot([cx, cx], [cy - 1.5 * scale, cy + 1.5 * scale], linewidth=1.0)
        ax.plot([cx - 0.9 * scale, cx + 1.3 * scale], [cy, cy], linewidth=1.0)
        _right_angle(cx, cy, 0.18 * scale)
        _label(cx, cy + 2.0 * scale, "Kite beginners' zone", size=11, weight="bold")
        _label(cx + 0.25 * scale, cy + 0.85 * scale, "10 m")
        _label(cx + 0.55 * scale, cy + 0.25 * scale, "6 m")
        _label(cx + 0.85 * scale, cy + 0.8 * scale, "7 m")
        _label(cx + 0.85 * scale, cy - 0.8 * scale, "7 m")
        _label(cx - 0.75 * scale, cy + 0.65 * scale, "5 m")
        _label(cx - 0.75 * scale, cy - 0.65 * scale, "5 m")
        _label(cx - 1.45 * scale, cy, "shared/internal\njoin 7 m", size=9)

    def _draw_right_triangle():
        ax.plot([0, 3.5], [0, 0], linewidth=1.8)
        ax.plot([0, 0], [0, 1.2], linewidth=1.8)
        ax.plot([0, 3.5], [1.2, 0], linewidth=1.8)
        _right_angle(0, 0, 0.25)
        _label(1.75, -0.25, "3.5 m")
        _label(-0.35, 0.6, "1.2 m")
        _label(1.9, 0.85, "Safety rail length = ? m")
        _label(1.75, 1.65, "Sloped bank cross-section", size=11, weight="bold")

    if "trapezium" in prompt or "trapezoid" in prompt:
        _draw_trapezium()
        ax.set_xlim(-3.2, 3.2)
        ax.set_ylim(-2.0, 2.2)
    elif "parallelogram" in prompt:
        _draw_parallelogram()
        ax.set_xlim(-3.2, 3.2)
        ax.set_ylim(-2.0, 2.2)
    elif "circle" in prompt or "bowl" in prompt or "radius" in prompt:
        _draw_circles()
        ax.set_xlim(-3.2, 3.2)
        ax.set_ylim(-2.0, 2.2)
    elif "kite" in prompt and ("composite" in prompt or "shared" in prompt or "join" in prompt):
        ax.plot([-3.0, -1.0, -1.0, -3.0, -3.0], [-1.0, -1.0, 1.0, 1.0, -1.0], linewidth=1.6)
        _label(-2.0, 1.35, "Chosen base\n(from Step 1)", size=10, weight="bold")
        _draw_kite(cx=0.6, cy=0, scale=0.8)
        ax.set_xlim(-3.6, 2.8)
        ax.set_ylim(-2.2, 2.4)
    elif "pythagoras" in prompt or "hypotenuse" in prompt or "slope" in prompt or "safety rail" in prompt:
        _draw_right_triangle()
        ax.set_xlim(-0.8, 4.2)
        ax.set_ylim(-0.7, 2.2)
    elif "design a" in prompt and "design b" in prompt:
        _draw_parallelogram(cx=-2.2, cy=0.2, scale=0.55, title_text="Design A complete")
        _draw_trapezium(cx=2.2, cy=0.2, scale=0.55, title_text="Design B complete")
        _label(-2.2, -1.55, "Bowl Option 1\nr = 3 m", size=9)
        _label(2.2, -1.55, "Bowl Option 2\nr = 4 m", size=9)
        ax.set_xlim(-4.2, 4.2)
        ax.set_ylim(-2.2, 2.0)
    else:
        _draw_parallelogram(cx=-1.8, cy=0, scale=0.55, title_text="Parallelogram")
        _draw_trapezium(cx=1.8, cy=0, scale=0.55, title_text="Trapezium")
        ax.set_xlim(-4.0, 4.0)
        ax.set_ylim(-2.0, 2.0)

    return _fig_to_png_bytes(fig)


# -----------------------------------------------------------------------------
# 2) Public API
# -----------------------------------------------------------------------------

# REPLACE your generate_diagram(...) with THIS patched version.
# ------------------------------------------------------------
# Assumes you've added the new functions earlier:
#   - render_image_gen_png_bytes(...)
#   - _image_gen_failed_png_bytes(...)
#   - CREDIT_COST_PER_IMAGE / set_credit_spend_fn(...)
#
# IMPORTANT: this keeps deterministic diagrams free,
# and makes image-gen paid + never-free (if spend fails, returns GEN FAILED png).

def generate_diagram(request: Dict[str, Any], user_ctx: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    user_ctx = user_ctx or {}
    prompt = (request.get("prompt") or "").strip()
    level = (request.get("level") or "J").strip().upper()
    
    user_ctx.setdefault("level", level)
    user_ctx.setdefault("tier", level_to_tier(level))

    subject = (request.get("subject") or "").strip().lower()
    archetype_hint = (request.get("archetype_hint") or "").strip()
    params = request.get("params") or {}

    # Optional fields (safe if not provided)
    where = (request.get("where") or request.get("title") or "").strip()
    notes = (request.get("notes") or "").strip()

    # When true, this function will EXECUTE image-gen (and charge credits)
    # When false (default), it will only return status="fallback" and let app decide.
    auto_image_gen = bool(request.get("auto_image_gen", False))

    archetype_id, confidence, why = resolve_archetype(prompt, archetype_hint, subject, level)

    def _fallback_payload(reason: str, archetype_id: Optional[str] = None, title: Optional[str] = None, debug_extra: Optional[Dict[str, Any]] = None):
        dbg = {"confidence": confidence, "why": why}
        if debug_extra:
            dbg.update(debug_extra)

        # If app wants library to auto-generate + charge, do it here.
        if auto_image_gen:
            png_bytes, bill = render_image_gen_png_bytes(
                where=where or (title or "Image"),
                prompt=prompt,
                notes=notes,
                user_ctx=user_ctx,
                reason="diagram_gen",
            )
            charged = bool(bill.get("charged"))
            cost = float(bill.get("cost", 0.0))
            gen_reason = bill.get("reason", "")

            return {
                "status": "ok" if charged else "error",
                "archetype_id": archetype_id,
                "title": title,
                "fallback": "image_gen",
                "cost_credits": cost,
                "reason": reason + (f" (image_gen: {gen_reason})" if gen_reason else ""),
                "debug": dbg,
                "mime": "image/png",
                "bytes": png_bytes,
                "billing": bill,
            }

        # Otherwise: just tell app “fallback needed” + show credit hint (no generation)
        return {
            "status": "fallback",
            "archetype_id": archetype_id,
            "title": title,
            "fallback": "image_gen",
            "cost_credits": _credit_cost_hint(user_ctx),
            "reason": reason,
            "debug": dbg,
            "mime": None,
            "bytes": None,
        }

    if not archetype_id:
        return _fallback_payload(
            reason="Could not confidently match a deterministic archetype.",
            archetype_id=None,
            title=None
        )

    meta = ARCHETYPES.get(archetype_id, {})
    title = meta.get("title")

    if not meta.get("deterministic", False):
        return _fallback_payload(
            reason="Archetype is marked non-deterministic in inventory.",
            archetype_id=archetype_id,
            title=title
        )

    renderer_name = meta.get("renderer")
    renderer_fn = globals().get(renderer_name) if renderer_name else None
    if not callable(renderer_fn):
        return _fallback_payload(
            reason="Deterministic archetype exists, but renderer not implemented yet.",
            archetype_id=archetype_id,
            title=title,
            debug_extra={"missing_renderer": renderer_name}
        )

    # Validate params before calling renderer
    merged_params = apply_defaults(archetype_id, params, level)
    merged_params.setdefault("tier", level_to_tier(level))

    user_ctx.setdefault("level", level)
    user_ctx.setdefault("tier", merged_params.get("tier"))

    # Infer missing params from the natural-language prompt (robust for LLM outputs)
    merged_params = _autofill_params_from_prompt(archetype_id, merged_params, prompt)
    
    ok, msg = validate_params(archetype_id, merged_params)

    # ---------------------------------------------------------------------
    # Stats Pack hardening: if data-driven archetypes lack required params,
    # automatically downgrade to blank templates (deterministic + free).
    # ---------------------------------------------------------------------
    if not ok:
        downgrade = {
            "histogram": "histogram_blank_axes",
            "stem_and_leaf": "stem_and_leaf_blank",
            "box_and_whisker": "box_and_whisker_blank",
            "scatter_plot": "scatter_plot_blank",
            "frequency_table": "frequency_table_blank",
            "two_way_table": "two_way_table_blank",
        }

        # Exam Pack hardening:
        # If these common exam diagram types are missing structured params,
        # supply safe defaults and continue deterministic rendering.
        if archetype_id == "venn_2":
            merged_params.setdefault("a_only", "")
            merged_params.setdefault("b_only", "")
            merged_params.setdefault("both", "")
            ok, msg = validate_params(archetype_id, merged_params)

        elif archetype_id == "triangle_diagram":
            merged_params.setdefault("labels", {})
            ok, msg = validate_params(archetype_id, merged_params)

        elif archetype_id == "quadratic_plot":
            merged_params.setdefault("a", 1)
            merged_params.setdefault("b", 0)
            merged_params.setdefault("c", 0)
            ok, msg = validate_params(archetype_id, merged_params)

        elif archetype_id == "linear_function_plot":
            merged_params.setdefault("m", 1)
            merged_params.setdefault("b", 0)
            ok, msg = validate_params(archetype_id, merged_params)

        elif archetype_id == "line_graph":
            merged_params.setdefault("x", [-5, 5])
            merged_params.setdefault("y", [0, 0])
            merged_params.setdefault("show_grid", True)
            ok, msg = validate_params(archetype_id, merged_params)

        elif archetype_id == "coordinate_plane_points":
            merged_params.setdefault("points", [])
            ok, msg = validate_params(archetype_id, merged_params)

        if not ok:
            blank_id = downgrade.get(archetype_id)

            if blank_id:
                archetype_id = blank_id
                meta = ARCHETYPES.get(archetype_id, {})
                title = meta.get("title") or title

                merged_params = apply_defaults(archetype_id, merged_params, level)
                merged_params.setdefault("tier", level_to_tier(level))

                ok, msg = validate_params(archetype_id, merged_params)
                if not ok:
                    return {
                        "status": "error",
                        "archetype_id": archetype_id,
                        "title": title,
                        "cost_credits": 0,
                        "message": msg,
                        "debug": {"confidence": confidence, "why": why, "params": merged_params},
                        "mime": None,
                        "bytes": None,
                    }
            else:
                return {
                    "status": "error",
                    "archetype_id": archetype_id,
                    "title": title,
                    "cost_credits": 0,
                    "message": msg,
                    "debug": {"confidence": confidence, "why": why, "params": merged_params},
                    "mime": None,
                    "bytes": None,
                }

            return {
                "status": "error",
                "archetype_id": archetype_id,
                "title": title,
                "cost_credits": 0,
                "message": msg,
                "debug": {"confidence": confidence, "why": why, "params": merged_params},
                "mime": None,
                "bytes": None,
            }

    # ---------------------------------------------------------------------
    # Render (deterministic)
    # ---------------------------------------------------------------------
    renderer_name = meta.get("renderer")
    renderer_fn = globals().get(renderer_name) if renderer_name else None
    if not callable(renderer_fn):
        return _fallback_payload(
            reason="Deterministic archetype exists, but renderer not implemented yet.",
            archetype_id=archetype_id,
            title=title,
            debug_extra={"missing_renderer": renderer_name}
        )

    try:
        png_bytes = renderer_fn(merged_params, level=level, subject=subject, title=title)
        return {
            "status": "ok",
            "archetype_id": archetype_id,
            "title": title,
            "cost_credits": 0,
            "mime": "image/png",
            "bytes": png_bytes,
            "debug": {"confidence": confidence, "why": why, "params": merged_params},
        }
    except Exception as e:
        return {
            "status": "error",
            "archetype_id": archetype_id,
            "title": title,
            "cost_credits": 0,
            "message": f"Renderer failed: {type(e).__name__}: {e}",
            "debug": {"confidence": confidence, "why": why, "params": merged_params},
            "mime": None,
            "bytes": None,
        }


def _bytes_to_tmp_png_path(png_bytes: bytes) -> str:
    """
    Write PNG bytes to a real temp file and return its path.
    (exports.py and pandoc expect file paths)
    """
    f = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    with open(f.name, "wb") as out:
        out.write(png_bytes or b"")
    return f.name


def render_diagram_png(where: str, prompt: str, notes: str = "", subtype: str = "") -> str:
    """
    Backwards-compatible API for exports.py
    Always returns a valid PNG *file path*.
    Deterministic remains free.
    If non-deterministic -> runs paid image-gen (credits must be wired) because export_confirm already happened.
    """
    req = {
        "where": (where or "").strip(),
        "prompt": (prompt or "").strip(),
        "notes": (notes or "").strip(),
        "level": "J",
        "subject": "",
        "archetype_hint": (subtype or "").strip(),
        "params": {},
        "auto_image_gen": True,  # IMPORTANT: actually generate+charge during export
    }

    res = generate_diagram(req, user_ctx=None)

    png_bytes = res.get("bytes")
    if isinstance(png_bytes, (bytes, bytearray)) and len(png_bytes) > 0:
        return _bytes_to_tmp_png_path(png_bytes)

    # Deterministic renderer failed (or unexpected response) -> never crash export
    err = res.get("message") or res.get("reason") or f"status={res.get('status')}"
    fallback_png = _image_gen_failed_png_bytes(where or "Diagram", f"Diagram render failed: {err}", prompt)
    return _bytes_to_tmp_png_path(fallback_png)


def render_image_gen_png(where: str, prompt: str, notes: str = "") -> str:
    """
    Backwards-compatible API for exports.py
    Always returns a valid PNG *file path*.
    Paid generation + credit spend happens inside render_image_gen_png_bytes().
    """
    png_bytes, _bill = render_image_gen_png_bytes(
        where=(where or "").strip(),
        prompt=(prompt or "").strip(),
        notes=(notes or "").strip(),
        user_ctx=None,
    )
    return _bytes_to_tmp_png_path(png_bytes)


# -----------------------------------------------------------------------------
# 3) Resolver
# -----------------------------------------------------------------------------
# Goal: allow a STRICT subject filter so that if the app passes subject="math",
# the resolver primarily considers only math archetypes, preventing cross-subject drift.
#
# Behavior:
#   - If subject is provided, we first search within matching-subject archetypes only.
#   - If no strong match is found there, we do a second pass across ALL archetypes.
#   - This keeps the system robust while still allowing fallback if subject data is weak.

def resolve_archetype(prompt: str, archetype_hint: str, subject: str, level: str) -> Tuple[Optional[str], float, str]:
    if archetype_hint and archetype_hint in ARCHETYPES:
        return archetype_hint, 1.0, "Used archetype_hint."

    subject = (subject or "").strip().lower()
    level = (level or "J").strip().upper()
    text = (prompt or "").lower()

    # --- Special routing: coordinate-plane line graph prompts ---
    if (
        ("coordinate grid" in text or "coordinate-plane" in text or "coordinate plane" in text)
        and ("straight line" in text or "line" in text or "gradient" in text or "slope" in text)
    ):
        return "line_graph", 0.95, "Matched coordinate-plane line graph phrasing."

    # --- Special routing: multiple-choice line graph option panels ---
    if (
        ("labelled a" in text or "labeled a" in text or "labelled a, b" in text or "labeled a, b" in text)
        and ("line" in text or "slope" in text or "gradient" in text)
        and ("negative" in text or "positive" in text or "horizontal" in text or "vertical" in text)
    ):
        return "line_graph", 0.95, "Matched line-graph multiple-choice panel phrasing."

    # --- Special routing: exam graph-matching line graphs ---
    if (
        ("match" in text or "matching" in text)
        and ("equation" in text or "equations" in text)
        and ("graph" in text or "graphs" in text)
        and ("line" in text or "linear" in text or "coordinate" in text)
    ):
        return "line_graph", 0.95, "Matched exam graph-matching line graph phrasing."
    
    # --- Special routing: exam triangle diagrams ---
    if (
        "triangle" in text
        and (
            "side length" in text
            or "side lengths" in text
            or "labelled" in text
            or "labeled" in text
            or "perpendicular height" in text
            or "height" in text
            or "base" in text
        )
    ):
        return "triangle_diagram", 0.95, "Matched exam triangle diagram phrasing."
    
    # --- Special routing: blank coordinate / Cartesian grid template ---
    if (
        ("blank" in text or "template" in text or "suitable for sketching" in text)
        and (
            "cartesian grid" in text
            or "coordinate grid" in text
            or "cartesian plane" in text
            or "coordinate plane" in text
            or "axes labelled x and y" in text
            or "axes labeled x and y" in text
        )
    ):
        return "coordinate_plane_points", 0.95, "Matched blank coordinate/Cartesian grid phrasing."

    # Blank stem-and-leaf template routing
    if ("stem" in text and "leaf" in text) and ("blank" in text or "template" in text or "provide space" in text):
        return "stem_and_leaf_blank", 0.95, "Matched blank stem-and-leaf template phrasing."

    # --- Special routing: blank box-and-whisker template ---
    if ("box" in text and "whisk" in text) and ("blank" in text or "template" in text or "provide space" in text or "use the axis below" in text):
        return "box_and_whisker_blank", 0.95, "Matched blank box-and-whisker template phrasing."

    # --- Special routing: blank scatter plot template ---
    if ("scatter" in text) and ("blank" in text or "template" in text or "use the grid below" in text or "provide space" in text):
        return "scatter_plot_blank", 0.95, "Matched blank scatter plot template phrasing."

    # --- Special routing: blank frequency table template ---
    if (("frequency table" in text) or ("tally table" in text)) and ("blank" in text or "template" in text or "fill in" in text or "provide space" in text):
        return "frequency_table_blank", 0.95, "Matched blank frequency/tally table template phrasing."

    # --- Special routing: blank two-way table template ---
    if (("two-way table" in text) or ("two way table" in text) or ("contingency table" in text) or ("cross tabulation" in text)) and ("blank" in text or "template" in text or "fill in" in text or "provide space" in text):
        return "two_way_table_blank", 0.95, "Matched blank two-way/contingency table template phrasing."

    
    # --- Special routing: blank histogram axes template (no bars) ---
    blank_axes_phrases = [
        "blank set of axes",
        "blank axes",
        "provide space for students",
        "space for students to draw",
        "use the axes below",
        "draw a histogram",
        "construct a histogram"
    ]
    if any(ph in text for ph in blank_axes_phrases):
        return "histogram_blank_axes", 0.95, "Matched blank histogram axes phrasing."


    def score_one(aid: str, meta: Dict[str, Any]) -> float:
        score = 0.0

        for s in (meta.get("synonyms", []) or []):
            s_l = s.lower().strip()
            if not s_l:
                continue
            if s_l in text:
                score += 1.0
            else:
                score += 0.15 * _token_overlap(s_l, text)

        if subject and subject in (meta.get("subjects", []) or []):
            score += 0.2
        if level and level in (meta.get("levels", []) or []):
            score += 0.1

        # deterministic preference / fallback penalty
        if meta.get("deterministic", False):
            score += 0.12
        else:
            score -= 0.10

        return score

    def best_match(candidate_ids) -> Tuple[Optional[str], float]:
        best_id_local = None
        best_score_local = float("-inf")
        for aid in candidate_ids:
            meta = ARCHETYPES.get(aid, {})
            sc = score_one(aid, meta)
            if sc > best_score_local:
                best_score_local = sc
                best_id_local = aid
        return best_id_local, best_score_local

    all_ids = list(ARCHETYPES.keys())

    # Pass 1: strict subject set (if subject is provided)
    if subject:
        subj_ids = [
            aid for aid, meta in ARCHETYPES.items()
            if subject in (meta.get("subjects", []) or [])
        ]
        if subj_ids:
            bid, bscore = best_match(subj_ids)
            if bscore >= 1.05:
                conf = min(1.0, 0.65 + 0.15 * (bscore - 1.05))
                return bid, conf, f"Strict subject pass matched '{subject}' with score={bscore:.2f}."

    # Pass 2: global (fallback to all archetypes)
    bid2, bscore2 = best_match(all_ids)
    if bscore2 >= 1.05:
        conf2 = min(1.0, 0.65 + 0.15 * (bscore2 - 1.05))
        return bid2, conf2, f"Global pass matched with score={bscore2:.2f}."

    return None, 0.0, "No strong match in subject pass or global pass; unresolved -> fallback."


# 2) OPTIONAL: Update debug_rank_archetypes to support strict_subject filtering too.
#    (Only if you added debug_rank_archetypes in Step 6.)

def debug_rank_archetypes(prompt: str, subject: str = "", level: str = "J", top_n: int = 8, strict_subject: bool = True) -> Dict[str, Any]:
    text = (prompt or "").lower()
    subject = (subject or "").strip().lower()
    level = (level or "J").strip().upper()

    def score_one(meta: Dict[str, Any]) -> float:
        score = 0.0
        for s in (meta.get("synonyms", []) or []):
            s_l = s.lower().strip()
            if not s_l:
                continue
            if s_l in text:
                score += 1.0
            else:
                score += 0.15 * _token_overlap(s_l, text)

        if subject and subject in (meta.get("subjects", []) or []):
            score += 0.2
        if level and level in (meta.get("levels", []) or []):
            score += 0.1

        if meta.get("deterministic", False):
            score += 0.12
        else:
            score -= 0.10

        return score

    candidates = list(ARCHETYPES.items())
    if strict_subject and subject:
        strict = [(aid, meta) for aid, meta in candidates if subject in (meta.get("subjects", []) or [])]
        if strict:
            candidates = strict

    scored = []
    for aid, meta in candidates:
        scored.append((aid, score_one(meta), meta.get("deterministic", False), meta.get("title", "")))

    scored.sort(key=lambda x: x[1], reverse=True)
    top = scored[:max(1, int(top_n))]

    return {
        "prompt": prompt,
        "subject": subject,
        "level": level,
        "strict_subject": strict_subject,
        "top": [
            {"archetype_id": aid, "score": round(score, 3), "deterministic": det, "title": ttl}
            for (aid, score, det, ttl) in top
        ],
    }


def _token_overlap(phrase: str, text: str) -> float:
    p = set(re.findall(r"[a-z0-9]+", phrase))
    t = set(re.findall(r"[a-z0-9]+", text))
    if not p or not t:
        return 0.0
    return len(p & t) / len(p)


# -----------------------------------------------------------------------------
# 4) Defaults + Validation
# -----------------------------------------------------------------------------

def apply_defaults(archetype_id: str, params: Dict[str, Any], level: str) -> Dict[str, Any]:
    meta = ARCHETYPES.get(archetype_id, {})
    defaults_by_level = meta.get("defaults_by_level", {}) or {}
    lvl_defaults = defaults_by_level.get(level, {}) or {}
    merged = dict(lvl_defaults)
    merged.update(params or {})
    return merged


def validate_params(archetype_id: str, params: Dict[str, Any]) -> Tuple[bool, str]:
    meta = ARCHETYPES.get(archetype_id, {})
    for k in (meta.get("required_params", []) or []):
        if k not in params:
            return False, f"Missing required param '{k}' for archetype '{archetype_id}'."
    return True, "ok"

# -----------------------------------------------------------------------------
# 4B) Prompt -> Params autofill (makes deterministic renderers robust)
# -----------------------------------------------------------------------------

def _parse_ints(text: str) -> List[int]:
    return [int(x) for x in re.findall(r"-?\d+", text or "")]


def _infer_stem_and_leaf_data(prompt: str) -> Optional[List[int]]:
    """
    Accepts either:
      - classic stem/leaf lines: 1 | 2 4 5 7 9
      - a plain data list: 12, 14, 15, 17, 19, 20 ...
    Returns: data as integers (e.g., 12, 14, 15...)
    """
    p = prompt or ""

    # Mode A: parse "stem | leaves"
    lines = [ln.strip() for ln in p.splitlines() if "|" in ln]
    data: List[int] = []
    got_any = False

    for ln in lines:
        m = re.match(r"^\s*(-?\d+)\s*\|\s*(.+?)\s*$", ln)
        if not m:
            continue
        stem = int(m.group(1))
        leaves_part = m.group(2).strip()
        # leaves are usually digits separated by spaces (sometimes commas)
        leaf_tokens = re.findall(r"-?\d+", leaves_part)
        if not leaf_tokens:
            continue
        got_any = True
        for tok in leaf_tokens:
            leaf = int(tok)
            # If leaf is 0-9, assume tens stem + ones leaf
            if 0 <= leaf <= 9:
                data.append(stem * 10 + leaf)
            else:
                # If leaves are multi-digit, just combine as string (fallback)
                try:
                    data.append(int(f"{stem}{abs(leaf)}"))
                except Exception:
                    pass

    if got_any and data:
        data.sort()
        return data

    # Mode B: fallback to extracting raw integers (works if prompt lists the dataset)
    ints = _parse_ints(p)
    if len(ints) >= 5:
        return sorted(ints)

    return None


def _infer_histogram_bins_freqs(prompt: str) -> Tuple[Optional[List[Any]], Optional[List[int]]]:
    """
    Tries to infer bins + frequencies from natural language, supporting:
      - "0-10: 2, 10-20: 6, 20-30: 8"
      - "Intervals: 0-10, 10-20, 20-30; Frequencies: 2, 6, 8"
    Returns bins in LABEL mode: ["0-10","10-20",...] and freqs [2,6,...]
    """
    p = prompt or ""

    # Mode A: explicit pairs like "0-10: 2"
    pair_pat = re.findall(r"(-?\d+(?:\.\d+)?)\s*[-–]\s*(-?\d+(?:\.\d+)?)\s*[:=]\s*(\d+)", p)
    if pair_pat:
        bins = [f"{a}-{b}" for (a, b, _) in pair_pat]
        freqs = [int(f) for (_, _, f) in pair_pat]
        if len(bins) == len(freqs) and len(bins) >= 2:
            return bins, freqs

    # Mode B: intervals listed, then frequencies listed
    intervals = re.findall(r"(-?\d+(?:\.\d+)?)\s*[-–]\s*(-?\d+(?:\.\d+)?)", p)
    freq_match = re.search(
        r"frequenc(?:y|ies)\s*(?:[:=\-])?\s*([0-9,\s]+)",
        p,
        flags=re.IGNORECASE
    )

    if intervals and freq_match:
        bins = [f"{a}-{b}" for (a, b) in intervals]
        freqs = [int(x) for x in re.findall(r"\d+", freq_match.group(1))]
        # Align lengths if possible
        if len(freqs) == len(bins) and len(bins) >= 2:
            return bins, freqs

    return None, None


def _autofill_params_from_prompt(archetype_id: str, params: Dict[str, Any], prompt: str) -> Dict[str, Any]:
    """
    If params are missing required keys, try to infer them from the prompt.
    Only fills missing keys; never overwrites user-provided params.
    """
    merged = dict(params or {})
    p = prompt or ""

    if archetype_id == "stem_and_leaf" and "data" not in merged:
        data = _infer_stem_and_leaf_data(p)
        if data:
            merged["data"] = data

    if archetype_id == "histogram":
        need_bins = "bins" not in merged
        need_freqs = "frequencies" not in merged
        if need_bins or need_freqs:
            bins, freqs = _infer_histogram_bins_freqs(p)
            if bins and freqs:
                merged.setdefault("bins", bins)
                merged.setdefault("frequencies", freqs)

    return merged


def _credit_cost_hint(user_ctx: Optional[Dict[str, Any]]) -> int:
    user_ctx = user_ctx or {}
    return int(user_ctx.get("image_gen_credit_cost", 1))


def level_to_tier(level: str) -> int:
    """
    Maps education level to complexity tier.
      P (Primary)    -> 1 (very simple)
      J (Junior)     -> 2 (moderate)
      S (Senior)     -> 3 (advanced)
      U (University) -> 4 (technical)
    """
    lv = (level or "J").strip().upper()
    return {"P": 1, "J": 2, "S": 3, "U": 4}.get(lv, 2)


# -----------------------------------------------------------------------------
# 5) Rendering helpers
# -----------------------------------------------------------------------------

def _fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# -----------------------------------------------------------------------------
# 6) Renderer Implementations (currently implemented)
# -----------------------------------------------------------------------------

def render_histogram_blank_axes(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Blank histogram axes template (no bars).
    If bins are provided, use them as x tick labels.
    If frequencies are provided, we use max(freq)+1 to set a sensible y-axis range.
    """
    bins = params.get("bins")  # optional
    freqs = params.get("frequencies")  # optional

    plot_title = params.get("title") or title or "Histogram"
    x_label = params.get("x_label") or ""
    y_label = params.get("y_label") or "Frequency"
    show_grid = bool(params.get("show_grid", True))

    # Determine y max
    y_max = params.get("y_max")
    if y_max is None:
        if isinstance(freqs, list) and freqs and all(isinstance(x, (int, float)) for x in freqs):
            y_max = int(max(freqs)) + 1
        else:
            y_max = 10  # safe default for Year 8 style tasks

    fig = plt.figure(figsize=(7.5, 4.8), dpi=220)
    ax = fig.add_axes([0.10, 0.18, 0.85, 0.72])

    ax.set_title(plot_title)
    ax.set_ylabel(y_label)
    if x_label:
        ax.set_xlabel(x_label)

    ax.set_ylim(0, max(1, int(y_max)))

    # If bins are labels (e.g., "0-10"), we set categorical ticks.
    if isinstance(bins, list) and bins:
        # label mode
        if all(isinstance(b, str) for b in bins):
            ax.set_xticks(range(len(bins)))
            ax.set_xticklabels(bins, rotation=0)
            ax.set_xlim(-0.5, len(bins) - 0.5)

        # edge mode: [0,10,20,...] -> label ticks at midpoints
        elif all(isinstance(b, (int, float)) for b in bins) and len(bins) >= 2:
            mids = [(bins[i] + bins[i + 1]) / 2 for i in range(len(bins) - 1)]
            labels = [f"{bins[i]}-{bins[i+1]}" for i in range(len(bins) - 1)]
            ax.set_xticks(mids)
            ax.set_xticklabels(labels, rotation=0)
            ax.set_xlim(min(bins), max(bins))
    else:
        # generic axes if nothing provided
        ax.set_xlim(0, 10)

    if show_grid:
        ax.grid(True, axis="y", linestyle="--", alpha=0.35)

    return _fig_to_png_bytes(fig)

    
def render_histogram(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic histogram renderer.
    Required params:
      - bins: list of bin edges OR list of bin labels
      - frequencies: list of frequencies per bin (same length as bins-1 if edges, or same length as labels)
    Optional:
      - title, x_label, y_label, show_grid
    Supports two input modes:
      A) Bin edges: bins=[0,10,20,30] and frequencies=[3,5,2]
      B) Bin labels: bins=["0-10","10-20","20-30"] and frequencies=[3,5,2]
    """
    bins = params["bins"]
    freqs = params["frequencies"]

    plot_title = params.get("title") or title or "Histogram"
    x_label = params.get("x_label") or ""
    y_label = params.get("y_label") or "Frequency"
    show_grid = bool(params.get("show_grid", True))

    if not isinstance(bins, list) or not isinstance(freqs, list) or not bins or not freqs:
        raise ValueError("bins and frequencies must be non-empty lists.")

    # Determine if bins are numeric edges or string labels
    bins_are_edges = False
    if all(isinstance(b, (int, float)) for b in bins):
        bins_are_edges = True
    elif all(isinstance(b, str) for b in bins):
        bins_are_edges = False
    else:
        raise ValueError("bins must be either all numbers (edges) or all strings (labels).")

    if bins_are_edges:
        # edges length must be freq+1
        if len(bins) != len(freqs) + 1:
            raise ValueError("For numeric bin edges, len(bins) must equal len(frequencies)+1.")
        # widths from edges
        widths = [float(bins[i + 1]) - float(bins[i]) for i in range(len(freqs))]
        if any(w <= 0 for w in widths):
            raise ValueError("Bin edges must be strictly increasing.")
        lefts = [float(bins[i]) for i in range(len(freqs))]
        centers = [lefts[i] + widths[i] / 2 for i in range(len(freqs))]
        x_ticks = centers
        x_tick_labels = [f"{bins[i]}–{bins[i+1]}" for i in range(len(freqs))]
    else:
        # labels mode: same length
        if len(bins) != len(freqs):
            raise ValueError("For labeled bins, len(bins) must equal len(frequencies).")
        # uniform width bars
        widths = [0.9 for _ in freqs]
        lefts = [i - 0.45 for i in range(len(freqs))]
        x_ticks = list(range(len(freqs)))
        x_tick_labels = [str(b) for b in bins]

    # Make the plot
    fig, ax = plt.subplots(figsize=(6.6, 4.0))
    ax.bar(lefts, freqs, width=widths, align="edge")

    ax.set_title(plot_title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)

    ax.set_xticks(x_ticks)
    ax.set_xticklabels(x_tick_labels, rotation=0)

    if show_grid:
        ax.grid(True, axis="y", linestyle="--", linewidth=0.6)

    return _fig_to_png_bytes(fig)

def render_box_and_whisker(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic box-and-whisker renderer.
    Required params:
      - min, q1, median, q3, max   (numbers)
    Optional:
      - title
      - show_outliers (bool)
      - outliers (list of numbers)
      - x_label (string)
      - show_grid (bool)
    Notes:
      - This renders a single boxplot (one dataset).
      - We validate ordering: min <= q1 <= median <= q3 <= max
    """
    vmin = float(params["min"])
    q1 = float(params["q1"])
    med = float(params["median"])
    q3 = float(params["q3"])
    vmax = float(params["max"])

    plot_title = params.get("title") or title or "Box and whisker plot"
    show_outliers = bool(params.get("show_outliers", False))
    outliers = params.get("outliers", []) or []
    x_label = params.get("x_label") or ""
    show_grid = bool(params.get("show_grid", True))

    if not (vmin <= q1 <= med <= q3 <= vmax):
        raise ValueError("Must satisfy min <= q1 <= median <= q3 <= max.")

    # Prepare stats dict for matplotlib bxp
    stats = [{
        "label": x_label if x_label else "",
        "whislo": vmin,
        "q1": q1,
        "med": med,
        "q3": q3,
        "whishi": vmax,
        "fliers": [float(x) for x in outliers] if (show_outliers and outliers) else [],
    }]

    fig, ax = plt.subplots(figsize=(6.2, 2.8))
    ax.set_title(plot_title)

    ax.bxp(stats, showfliers=bool(show_outliers and outliers), vert=False)

    if show_grid:
        ax.grid(True, axis="x", linestyle="--", linewidth=0.6)

    # Keep y-axis clean for a single box
    ax.set_yticks([])
    if x_label:
        ax.set_xlabel(x_label)

    return _fig_to_png_bytes(fig)


def render_box_and_whisker_blank(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Blank box-and-whisker template:
    - number line axis (level-appropriate default range)
    - empty space above axis for students to draw box + whiskers
    """
    plot_title = (params.get("title") or title or "Box and whisker plot").strip()

    # Level-aware defaults come from apply_defaults(), but we hard-guard here too.
    x_min = params.get("x_min", 0)
    x_max = params.get("x_max", 40 if (level or "J").upper() == "J" else 100)
    x_step = params.get("x_step", 5 if (level or "J").upper() == "J" else 10)

    show_grid = bool(params.get("show_grid", True))
    show_key = bool(params.get("show_key", True))

    # Defensive numeric conversion
    try:
        x_min = float(x_min)
        x_max = float(x_max)
        x_step = float(x_step)
    except Exception:
        x_min, x_max, x_step = 0.0, 40.0, 5.0

    if x_max <= x_min:
        x_max = x_min + 10

    fig = plt.figure(figsize=(7.5, 4.6), dpi=220)
    ax = fig.add_axes([0.10, 0.20, 0.85, 0.70])

    ax.set_title(plot_title)

    # Axis range
    ax.set_xlim(x_min, x_max)
    ax.set_ylim(0, 1)

    # Clean template look
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["top"].set_visible(False)

    # Ticks
    if x_step > 0:
        ticks = []
        t = x_min
        # avoid float drift
        while t <= x_max + 1e-9:
            ticks.append(t)
            t += x_step
        ax.set_xticks(ticks)

    if show_grid:
        ax.grid(True, axis="x", linestyle="--", alpha=0.35)

    # Drawing area hint (light guide line)
    ax.hlines(0.55, x_min, x_max, linewidth=1.0, alpha=0.25)

    if show_key:
        ax.text(
            0.0, -0.28,
            "Students: draw the box from Q1 to Q3, a line at the median, and whiskers to min/max.",
            fontsize=10,
            transform=ax.transAxes
        )

    return _fig_to_png_bytes(fig)


def render_coordinate_plane_points(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic coordinate plane points renderer.
    Required params:
      - points: list of points, each point can be:
          {"x": 2, "y": 3, "label": "A"}  (label optional)
        or a tuple/list:
          (2, 3, "A") or (2, 3)
    Optional params:
      - title
      - x_min, x_max, y_min, y_max (numbers)
      - show_grid (bool)
      - x_label, y_label (strings)
      - show_origin_axes (bool)  # draw thicker x=0 and y=0
      - point_size (int)
    Behavior:
      - Auto-ranges if bounds not provided.
      - Expands bounds slightly for readability.
    """
    pts = params["points"]
    plot_title = params.get("title") or title or "Coordinate plane"
    show_grid = bool(params.get("show_grid", True))
    x_label = params.get("x_label") or "x"
    y_label = params.get("y_label") or "y"
    show_origin_axes = bool(params.get("show_origin_axes", True))
    point_size = int(params.get("point_size", 40))

    if not isinstance(pts, list) or not pts:
        raise ValueError("points must be a non-empty list.")

    xs, ys, labels = [], [], []
    for p in pts:
        if isinstance(p, dict):
            x = float(p.get("x"))
            y = float(p.get("y"))
            lab = str(p.get("label", "")).strip()
        elif isinstance(p, (list, tuple)):
            if len(p) < 2:
                raise ValueError("Point tuple/list must have at least (x, y).")
            x = float(p[0])
            y = float(p[1])
            lab = str(p[2]).strip() if len(p) >= 3 and p[2] is not None else ""
        else:
            raise ValueError("Each point must be a dict or tuple/list.")

        xs.append(x)
        ys.append(y)
        labels.append(lab)

    # Bounds: use provided or auto from data
    x_min = params.get("x_min", None)
    x_max = params.get("x_max", None)
    y_min = params.get("y_min", None)
    y_max = params.get("y_max", None)

    if x_min is None: x_min = min(xs)
    if x_max is None: x_max = max(xs)
    if y_min is None: y_min = min(ys)
    if y_max is None: y_max = max(ys)

    x_min = float(x_min)
    x_max = float(x_max)
    y_min = float(y_min)
    y_max = float(y_max)

    # Expand bounds slightly for readability (and avoid zero-span)
    if x_max == x_min:
        x_max += 1.0
        x_min -= 1.0
    if y_max == y_min:
        y_max += 1.0
        y_min -= 1.0

    pad_x = 0.15 * (x_max - x_min)
    pad_y = 0.15 * (y_max - y_min)
    x_min -= pad_x
    x_max += pad_x
    y_min -= pad_y
    y_max += pad_y

    fig, ax = plt.subplots(figsize=(6.0, 5.0))
    ax.set_title(plot_title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)

    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min, y_max)

    if show_grid:
        ax.grid(True, linestyle="--", linewidth=0.6)

    # Draw origin axes (x=0 and y=0) if within view
    if show_origin_axes:
        if x_min <= 0 <= x_max:
            ax.axvline(0, linewidth=1.2)
        if y_min <= 0 <= y_max:
            ax.axhline(0, linewidth=1.2)

    ax.scatter(xs, ys, s=point_size)

    # Labels
    for x, y, lab in zip(xs, ys, labels):
        if lab:
            ax.text(x, y, f" {lab}", va="bottom", ha="left", fontsize=10)

    return _fig_to_png_bytes(fig)
    

def render_frequency_table(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic frequency table renderer (as an image).
    Required params:
      - rows: list of rows. Each row can be:
          {"label": "Cats", "freq": 5}
        or
          {"value": "Cats", "frequency": 5}
    Optional params:
      - title
      - show_tally (bool)  # mostly for Primary
      - show_total (bool)
      - label_header (str)   default: "Value"
      - freq_header (str)    default: "Frequency"
      - tally_header (str)   default: "Tally"
    """
    rows = params["rows"]
    plot_title = params.get("title") or title or "Frequency table"

    show_tally = bool(params.get("show_tally", False))
    show_total = bool(params.get("show_total", True))

    label_header = params.get("label_header") or "Value"
    freq_header = params.get("freq_header") or "Frequency"
    tally_header = params.get("tally_header") or "Tally"

    if not isinstance(rows, list) or not rows:
        raise ValueError("rows must be a non-empty list.")

    parsed = []
    total = 0
    for r in rows:
        if not isinstance(r, dict):
            raise ValueError("Each row must be a dict like {'label':..., 'freq':...}.")
        label = r.get("label", r.get("value", r.get("category", "")))
        freq = r.get("freq", r.get("frequency", r.get("count", None)))
        if label is None or str(label).strip() == "":
            raise ValueError("Each row must have a non-empty label/value/category.")
        if not isinstance(freq, (int, float)):
            raise ValueError("Each row must have numeric freq/frequency/count.")
        freq_num = int(freq) if float(freq).is_integer() else float(freq)
        total += float(freq)
        parsed.append((str(label).strip(), freq_num))

    # Build table content
    headers = [label_header]
    if show_tally:
        headers.append(tally_header)
    headers.append(freq_header)

    cell_text = []
    for label, freq in parsed:
        row_cells = [label]
        if show_tally:
            row_cells.append(_tally_marks(int(freq)) if isinstance(freq, (int, float)) else "")
        row_cells.append(str(freq))
        cell_text.append(row_cells)

    if show_total:
        total_row = ["Total"]
        if show_tally:
            total_row.append("")
        total_row.append(str(int(total) if float(total).is_integer() else total))
        cell_text.append(total_row)

    # Render using matplotlib table
    fig_h = max(2.0, 0.55 + 0.35 * len(cell_text))
    fig, ax = plt.subplots(figsize=(6.2, fig_h))
    ax.set_title(plot_title)
    ax.axis("off")

    tbl = ax.table(
        cellText=cell_text,
        colLabels=headers,
        cellLoc="center",
        colLoc="center",
        loc="center",
    )

    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    tbl.scale(1.0, 1.25)

    return _fig_to_png_bytes(fig)


def _tally_marks(n: int) -> str:
    """
    Returns a simple tally string grouped in 5s: '||||/ ||||/'
    (Avoids special glyphs to keep rendering stable.)
    """
    if n <= 0:
        return ""
    groups = n // 5
    rem = n % 5
    parts = ["||||/" for _ in range(groups)]
    if rem:
        parts.append("|" * rem)
    return " ".join(parts)


def render_frequency_table_blank(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Blank frequency table template.
    Creates an empty grid with headers + a fixed number of empty rows.
    """
    plot_title = params.get("title") or title or "Frequency table"
    show_tally = bool(params.get("show_tally", False))
    show_total = bool(params.get("show_total", False))

    label_header = params.get("label_header") or "Category"
    freq_header = params.get("freq_header") or "Frequency"
    tally_header = params.get("tally_header") or "Tally"

    # How many blank rows?
    row_count = params.get("row_count", 8)
    try:
        row_count = int(row_count)
    except Exception:
        row_count = 8
    row_count = max(3, min(row_count, 20))

    headers = [label_header]
    if show_tally:
        headers.append(tally_header)
    headers.append(freq_header)

    cell_text = []
    for _ in range(row_count):
        row = [""]
        if show_tally:
            row.append("")
        row.append("")
        cell_text.append(row)

    if show_total:
        total_row = ["Total"]
        if show_tally:
            total_row.append("")
        total_row.append("")
        cell_text.append(total_row)

    fig_h = max(2.2, 0.8 + 0.35 * len(cell_text))
    fig, ax = plt.subplots(figsize=(6.2, fig_h))
    ax.set_title(plot_title)
    ax.axis("off")

    tbl = ax.table(
        cellText=cell_text,
        colLabels=headers,
        cellLoc="center",
        colLoc="center",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    tbl.scale(1.0, 1.35)

    return _fig_to_png_bytes(fig)


def render_two_way_table_blank(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Blank two-way table template.
    If row_labels/col_labels are provided, use them; otherwise generate Row 1..n / Col 1..n.
    """
    plot_title = params.get("title") or title or "Two-way table"
    show_totals = bool(params.get("show_totals", True))

    corner_label = params.get("corner_label") or ""
    row_total_label = params.get("row_total_label") or "Total"
    col_total_label = params.get("col_total_label") or "Total"

    # Labels or counts
    row_labels = params.get("row_labels")
    col_labels = params.get("col_labels")

    if not (isinstance(row_labels, list) and row_labels):
        rc = params.get("row_count", 2)
        try:
            rc = int(rc)
        except Exception:
            rc = 2
        rc = max(2, min(rc, 10))
        row_labels = [f"Row {i+1}" for i in range(rc)]

    if not (isinstance(col_labels, list) and col_labels):
        cc = params.get("col_count", 2)
        try:
            cc = int(cc)
        except Exception:
            cc = 2
        cc = max(2, min(cc, 10))
        col_labels = [f"Col {i+1}" for i in range(cc)]

    # Build headers
    headers = [corner_label] + [str(c) for c in col_labels]
    if show_totals:
        headers.append(col_total_label)

    # Build rows (blank cells)
    cell_text = []
    for r in row_labels:
        row = [str(r)] + [""] * len(col_labels)
        if show_totals:
            row.append("")
        cell_text.append(row)

    # Optional bottom totals row (blank)
    if show_totals:
        total_row = [row_total_label] + [""] * len(col_labels) + [""]
        cell_text.append(total_row)

    fig_h = max(2.2, 0.9 + 0.35 * len(cell_text))
    fig_w = max(6.2, 0.9 + 1.05 * len(headers))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_title(plot_title)
    ax.axis("off")

    tbl = ax.table(
        cellText=cell_text,
        colLabels=headers,
        cellLoc="center",
        colLoc="center",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    tbl.scale(1.0, 1.35)

    return _fig_to_png_bytes(fig)


def render_two_way_table(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic two-way table renderer (as an image).
    Required params:
      - row_labels: list[str]
      - col_labels: list[str]
      - values_matrix: list[list[number]] with shape (len(row_labels), len(col_labels))
    Optional params:
      - title
      - show_totals (bool) default True
      - corner_label (str) default "" (top-left corner label)
      - row_total_label (str) default "Total"
      - col_total_label (str) default "Total"
    """
    row_labels = params["row_labels"]
    col_labels = params["col_labels"]
    matrix = params["values_matrix"]

    plot_title = params.get("title") or title or "Two-way table"
    show_totals = bool(params.get("show_totals", True))

    corner_label = params.get("corner_label") or ""
    row_total_label = params.get("row_total_label") or "Total"
    col_total_label = params.get("col_total_label") or "Total"

    if not (isinstance(row_labels, list) and row_labels):
        raise ValueError("row_labels must be a non-empty list.")
    if not (isinstance(col_labels, list) and col_labels):
        raise ValueError("col_labels must be a non-empty list.")
    if not (isinstance(matrix, list) and len(matrix) == len(row_labels)):
        raise ValueError("values_matrix must be a list with one row per row_label.")

    # Validate matrix shape
    for r in matrix:
        if not (isinstance(r, list) and len(r) == len(col_labels)):
            raise ValueError("Each row of values_matrix must match length of col_labels.")
        for v in r:
            if not isinstance(v, (int, float)):
                raise ValueError("All values in values_matrix must be numeric.")

    # Prepare display grid including headers
    # We'll build a full cellText with row header column included.
    display_col_labels = [corner_label] + [str(c) for c in col_labels]
    display_rows = []

    row_totals = [sum(float(v) for v in row) for row in matrix]
    col_totals = [sum(float(matrix[r][c]) for r in range(len(row_labels))) for c in range(len(col_labels))]
    grand_total = sum(row_totals)

    for i, rlab in enumerate(row_labels):
        row_vals = [str(rlab)]
        row_vals += [str(int(v) if float(v).is_integer() else v) for v in matrix[i]]
        display_rows.append(row_vals)

    if show_totals:
        # append row totals column
        display_col_labels.append(col_total_label)
        for i in range(len(display_rows)):
            rt = row_totals[i]
            display_rows[i].append(str(int(rt) if float(rt).is_integer() else rt))

        # append totals row at bottom
        totals_row = [row_total_label]
        totals_row += [str(int(ct) if float(ct).is_integer() else ct) for ct in col_totals]
        totals_row.append(str(int(grand_total) if float(grand_total).is_integer() else grand_total))
        display_rows.append(totals_row)

    # Render
    fig_h = max(2.2, 0.70 + 0.35 * len(display_rows))
    fig_w = max(6.2, 1.2 + 0.9 * len(display_col_labels))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_title(plot_title)
    ax.axis("off")

    tbl = ax.table(
        cellText=display_rows,
        colLabels=display_col_labels,
        cellLoc="center",
        colLoc="center",
        loc="center",
    )

    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    tbl.scale(1.0, 1.25)

    return _fig_to_png_bytes(fig)


def render_coordinate_plane_points(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic coordinate grid renderer.
    Supports blank grids and optional plotted points.
    """
    plot_title = params.get("title") or title or "Coordinate plane"

    x_min = int(params.get("x_min", -5))
    x_max = int(params.get("x_max", 5))
    y_min = int(params.get("y_min", -5))
    y_max = int(params.get("y_max", 5))
    points = params.get("points") or []

    fig, ax = plt.subplots(figsize=(5.2, 5.2))
    ax.set_title(plot_title)
    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min, y_max)
    ax.set_aspect("equal", adjustable="box")

    ax.axhline(0, linewidth=1.2)
    ax.axvline(0, linewidth=1.2)

    ax.set_xticks(range(x_min, x_max + 1))
    ax.set_yticks(range(y_min, y_max + 1))
    ax.grid(True, linewidth=0.5, alpha=0.45)

    for p in points:
        try:
            if isinstance(p, dict):
                x = float(p.get("x"))
                y = float(p.get("y"))
                label = str(p.get("label", "")).strip()
            elif isinstance(p, (list, tuple)) and len(p) >= 2:
                x = float(p[0])
                y = float(p[1])
                label = str(p[2]).strip() if len(p) >= 3 else ""
            else:
                continue

            ax.plot([x], [y], marker="o")
            if label:
                ax.text(x + 0.15, y + 0.15, label, fontsize=9)
        except Exception:
            continue

    return _fig_to_png_bytes(fig)


def render_linear_function_plot(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic straight-line graph renderer.
    """
    plot_title = params.get("title") or title or "Linear function"

    m = float(params.get("m", 1))
    b = float(params.get("b", 0))
    x_min = int(params.get("x_min", -5))
    x_max = int(params.get("x_max", 5))
    y_min = int(params.get("y_min", -5))
    y_max = int(params.get("y_max", 5))

    xs = [x_min + i * (x_max - x_min) / 100 for i in range(101)]
    ys = [m * x + b for x in xs]

    fig, ax = plt.subplots(figsize=(5.2, 5.2))
    ax.set_title(plot_title)
    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min, y_max)
    ax.set_aspect("equal", adjustable="box")

    ax.axhline(0, linewidth=1.2)
    ax.axvline(0, linewidth=1.2)
    ax.set_xticks(range(x_min, x_max + 1))
    ax.set_yticks(range(y_min, y_max + 1))
    ax.grid(True, linewidth=0.5, alpha=0.45)
    ax.plot(xs, ys, linewidth=1.8)

    return _fig_to_png_bytes(fig)


def render_quadratic_plot(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic quadratic/parabola renderer.
    """
    plot_title = params.get("title") or title or "Quadratic function"

    a = float(params.get("a", 1))
    b = float(params.get("b", 0))
    c = float(params.get("c", 0))
    x_min = int(params.get("x_min", -5))
    x_max = int(params.get("x_max", 5))
    y_min = int(params.get("y_min", -6))
    y_max = int(params.get("y_max", 10))

    xs = [x_min + i * (x_max - x_min) / 200 for i in range(201)]
    ys = [a * x * x + b * x + c for x in xs]

    fig, ax = plt.subplots(figsize=(5.2, 5.2))
    ax.set_title(plot_title)
    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min, y_max)
    ax.set_aspect("equal", adjustable="box")

    ax.axhline(0, linewidth=1.2)
    ax.axvline(0, linewidth=1.2)
    ax.set_xticks(range(x_min, x_max + 1))
    ax.set_yticks(range(y_min, y_max + 1))
    ax.grid(True, linewidth=0.5, alpha=0.45)
    ax.plot(xs, ys, linewidth=1.8)

    return _fig_to_png_bytes(fig)


def render_venn_2(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic two-circle Venn diagram renderer.
    Works as a blank template or with region values if provided.
    """
    plot_title = params.get("title") or title or "Venn diagram"
    label_a = params.get("label_a") or "A"
    label_b = params.get("label_b") or "B"

    fig, ax = plt.subplots(figsize=(5.8, 3.8))
    ax.set_title(plot_title)
    ax.set_aspect("equal")
    ax.axis("off")

    rect = plt.Rectangle((-3.0, -1.7), 6.0, 3.1, fill=False, linewidth=1.2)
    ax.add_patch(rect)

    c1 = plt.Circle((-0.8, 0), 1.15, fill=False, linewidth=1.6)
    c2 = plt.Circle((0.8, 0), 1.15, fill=False, linewidth=1.6)
    ax.add_patch(c1)
    ax.add_patch(c2)

    ax.text(-1.35, 1.28, str(label_a), ha="center", va="center", fontsize=11)
    ax.text(1.35, 1.28, str(label_b), ha="center", va="center", fontsize=11)

    if "a_only" in params:
        ax.text(-1.25, 0, str(params.get("a_only", "")), ha="center", va="center", fontsize=11)
    if "both" in params:
        ax.text(0, 0, str(params.get("both", "")), ha="center", va="center", fontsize=11)
    if "b_only" in params:
        ax.text(1.25, 0, str(params.get("b_only", "")), ha="center", va="center", fontsize=11)
    if "neither" in params:
        ax.text(2.45, -1.25, str(params.get("neither", "")), ha="center", va="center", fontsize=10)

    ax.set_xlim(-3.2, 3.2)
    ax.set_ylim(-1.9, 1.7)

    return _fig_to_png_bytes(fig)


def render_triangle_diagram(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic labelled triangle renderer.
    """
    plot_title = params.get("title") or title or "Triangle diagram"
    show_right_angle = bool(params.get("show_right_angle", False))
    labels = params.get("labels") or {}

    fig, ax = plt.subplots(figsize=(5.8, 3.8))
    ax.set_title(plot_title)
    ax.set_aspect("equal")
    ax.axis("off")

    A = (0, 0)
    B = (4, 0)
    C = (1.2, 2.2)

    ax.plot([A[0], B[0], C[0], A[0]], [A[1], B[1], C[1], A[1]], linewidth=1.8)

    ax.text(A[0] - 0.18, A[1] - 0.18, labels.get("A", "A"), fontsize=11)
    ax.text(B[0] + 0.10, B[1] - 0.18, labels.get("B", "B"), fontsize=11)
    ax.text(C[0], C[1] + 0.18, labels.get("C", "C"), fontsize=11)

    ax.text(2.0, -0.28, str(labels.get("AB", labels.get("base", ""))), ha="center", fontsize=10)
    ax.text(2.75, 1.25, str(labels.get("BC", "")), ha="center", fontsize=10)
    ax.text(0.35, 1.15, str(labels.get("AC", "")), ha="center", fontsize=10)

    if show_right_angle:
        ax.plot([0, 0.35, 0.35], [0.35, 0.35, 0], linewidth=1.0)

    ax.set_xlim(-0.6, 4.6)
    ax.set_ylim(-0.6, 2.8)

    return _fig_to_png_bytes(fig)


def render_stem_and_leaf_blank(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Blank stem-and-leaf template (deterministic).
    Always returns PNG bytes.
    """
    plot_title = (params.get("title") or title or "Stem-and-leaf plot").strip()
    rows = params.get("rows", 12)
    stem_unit = params.get("stem_unit", 10)
    show_key = bool(params.get("show_key", True))

    try:
        rows = int(rows)
    except Exception:
        rows = 12
    rows = max(6, min(rows, 25))

    # Basic template text grid
    fig = plt.figure(figsize=(7.0, 6.0), dpi=220)
    ax = fig.add_axes([0.08, 0.10, 0.86, 0.82])
    ax.axis("off")
    ax.set_title(plot_title)

    # Column positions in axes coords
    x_stem = 0.15
    x_bar = 0.23
    x_leaf = 0.28

    # Header
    ax.text(x_stem, 0.92, "Stem", fontsize=12, weight="bold", transform=ax.transAxes)
    ax.text(x_bar, 0.92, "|", fontsize=12, weight="bold", transform=ax.transAxes)
    ax.text(x_leaf, 0.92, "Leaves", fontsize=12, weight="bold", transform=ax.transAxes)

    # Rows
    top = 0.88
    row_h = (top - 0.18) / rows
    for i in range(rows):
        y = top - i * row_h
        ax.text(x_stem, y, "", fontsize=12, transform=ax.transAxes)
        ax.text(x_bar, y, "|", fontsize=12, transform=ax.transAxes)
        # draw an underline area for leaves
        ax.plot([x_leaf, 0.92], [y - 0.01, y - 0.01], linewidth=1.0, alpha=0.7, transform=ax.transAxes)

    # Optional key
    if show_key:
        ax.text(0.10, 0.08, f"Key: 3 | 5 means {3*int(stem_unit)} + 5 = {3*int(stem_unit)+5}", fontsize=10, transform=ax.transAxes)

    return _fig_to_png_bytes(fig)
    

def render_stem_and_leaf(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic stem-and-leaf plot.
    Required params:
      - data: list of integers (or floats that are whole numbers)
    Optional params:
      - title
      - stem_unit (int) default 10   # stem = value // stem_unit, leaf = value % stem_unit
      - sort_leaves (bool) default True
      - show_key (bool) default True
      - key_example (str) optional override (e.g., "3 | 7 means 37")
    Notes:
      - For typical school use, stem_unit=10 gives stems as tens and leaves as ones.
      - Data should be non-negative for clean display (we allow negatives but show stems accordingly).
    """
    data = params["data"]
    plot_title = params.get("title") or title or "Stem-and-leaf plot"
    stem_unit = int(params.get("stem_unit", 10))
    sort_leaves = bool(params.get("sort_leaves", True))
    show_key = bool(params.get("show_key", True))
    key_example = params.get("key_example")

    if not isinstance(data, list) or not data:
        raise ValueError("data must be a non-empty list.")
    if stem_unit <= 0:
        raise ValueError("stem_unit must be a positive integer.")

    # Normalize to ints if whole
    vals = []
    for v in data:
        if not isinstance(v, (int, float)):
            raise ValueError("All data values must be numeric.")
        vf = float(v)
        if not vf.is_integer():
            raise ValueError("Stem-and-leaf expects whole numbers (integers).")
        vals.append(int(vf))

    # Build stems
    stems: Dict[int, list] = {}
    for v in vals:
        stem = int(math.floor(v / stem_unit))
        leaf = int(abs(v) % stem_unit)  # leaf as digit group
        stems.setdefault(stem, []).append(leaf)

    # Sort stems
    stem_keys = sorted(stems.keys())

    # Prepare rows: each row = [stem, leaves_str]
    rows = []
    for stem in stem_keys:
        leaves = stems[stem]
        if sort_leaves:
            leaves = sorted(leaves)
        leaves_str = " ".join(str(l) for l in leaves)
        rows.append([str(stem), leaves_str])

    # Key
    if show_key:
        if key_example and isinstance(key_example, str) and key_example.strip():
            key_line = key_example.strip()
        else:
            # pick an example from data
            example = sorted(vals)[len(vals) // 2]
            ex_stem = int(math.floor(example / stem_unit))
            ex_leaf = int(abs(example) % stem_unit)
            # reconstruct meaning
            meaning = ex_stem * stem_unit + (ex_leaf if example >= 0 else -ex_leaf)
            key_line = f"{ex_stem} | {ex_leaf} means {meaning}"
    else:
        key_line = ""

    # Render as table
    fig_h = max(2.2, 0.8 + 0.35 * len(rows))
    fig, ax = plt.subplots(figsize=(6.2, fig_h))
    ax.set_title(plot_title)
    ax.axis("off")

    tbl = ax.table(
        cellText=rows,
        colLabels=["Stem", "Leaves"],
        cellLoc="left",
        colLoc="center",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)
    tbl.scale(1.0, 1.35)

    if key_line:
        ax.text(0.02, 0.02, f"Key: {key_line}", transform=ax.transAxes, fontsize=10, ha="left", va="bottom")

    return _fig_to_png_bytes(fig)


def render_stem_and_leaf_blank(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Blank stem-and-leaf worksheet template.
    Level-aware:
      - J: fewer rows / simpler layout
      - S/U: more rows / more writing space
    Optional params:
      - stems: explicit list of stem values (ints)
      - min/max: infer stems range from min/max and stem_unit
      - stem_unit: default 10 (tens as stems)
      - rows: number of template rows (writing lines)
      - show_key / key_example
    """
    lv = (level or "J").strip().upper()
    stem_unit = int(params.get("stem_unit", 10))
    rows = int(params.get("rows", 12 if lv == "J" else 18 if lv == "S" else 22))
    show_key = bool(params.get("show_key", True))
    key_example = (params.get("key_example") or "").strip()

    # Determine stems to display
    stems = params.get("stems")
    if isinstance(stems, list) and all(isinstance(x, (int, float)) for x in stems):
        stems_list = [int(x) for x in stems][:rows]
    else:
        mn = params.get("min")
        mx = params.get("max")
        stems_list = []

        if isinstance(mn, (int, float)) and isinstance(mx, (int, float)) and stem_unit > 0:
            s0 = int(mn) // stem_unit
            s1 = int(mx) // stem_unit
            stems_list = list(range(s0, s1 + 1))
        else:
            # Safe defaults by level (not childish for seniors)
            if lv == "J":
                stems_list = list(range(0, 10))      # 0–9
            elif lv == "S":
                stems_list = list(range(0, 16))      # 0–15
            else:
                stems_list = list(range(0, 20))      # 0–19

    # Trim/pad to rows
    stems_list = stems_list[:rows]
    while len(stems_list) < rows:
        stems_list.append("")

    plot_title = (params.get("title") or title or "Stem-and-leaf plot (blank)").strip()

    fig = plt.figure(figsize=(8.27, 11.69), dpi=220)
    ax = fig.add_axes([0.08, 0.08, 0.84, 0.84])
    ax.axis("off")

    # Title
    ax.text(0.0, 1.02, plot_title, fontsize=16, weight="bold", va="bottom", transform=ax.transAxes)

    # Layout constants
    top = 0.93
    bottom = 0.10 if show_key else 0.06
    left = 0.00
    mid = 0.22   # stem column width
    right = 0.98

    # Outer box
    ax.add_patch(plt.Rectangle((left, bottom), right-left, top-bottom, fill=False, linewidth=1.3, transform=ax.transAxes))

    # Vertical divider (stem | leaves)
    ax.plot([mid, mid], [bottom, top], linewidth=1.6, transform=ax.transAxes)

    # Header text
    ax.text(left + 0.02, top + 0.01, "Stem", fontsize=12, weight="bold", transform=ax.transAxes)
    ax.text(mid + 0.02, top + 0.01, "Leaves", fontsize=12, weight="bold", transform=ax.transAxes)

    # Row lines
    usable_h = (top - bottom)
    row_h = usable_h / max(1, rows)
    for i in range(rows + 1):
        y = top - i * row_h
        ax.plot([left, right], [y, y], linewidth=0.8, alpha=0.7, transform=ax.transAxes)

    # Stems text
    for i, s in enumerate(stems_list):
        y = top - (i + 0.65) * row_h
        ax.text(left + 0.05, y, str(s), fontsize=12, va="center", transform=ax.transAxes)

    # Key
    if show_key:
        if not key_example and stem_unit == 10:
            key_example = "3 | 7 means 37"
        elif not key_example:
            key_example = f"3 | 7 means 3×{stem_unit} + 7"

        ax.text(0.0, 0.04, f"Key: {key_example}", fontsize=11, transform=ax.transAxes)

    return _fig_to_png_bytes(fig)


def render_water_cycle(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic Water Cycle diagram.
    Optional params:
      - title
      - labels: dict to override default labels, e.g.:
          {
            "evaporation": "Evaporation",
            "condensation": "Condensation",
            "precipitation": "Precipitation",
            "collection": "Collection",
            "transpiration": "Transpiration"
          }
      - include_transpiration (bool) default True
      - include_runoff (bool) default True
      - include_infiltration (bool) default False
      - show_sun (bool) default True
    """
    plot_title = params.get("title") or title or "The Water Cycle"

    labels = params.get("labels") or {}
    if not isinstance(labels, dict):
        labels = {}

    lab_evap = labels.get("evaporation", "Evaporation")
    lab_cond = labels.get("condensation", "Condensation")
    lab_prec = labels.get("precipitation", "Precipitation")
    lab_coll = labels.get("collection", "Collection")
    lab_trans = labels.get("transpiration", "Transpiration")
    lab_runoff = labels.get("runoff", "Runoff")
    lab_infil = labels.get("infiltration", "Infiltration")

    include_transpiration = bool(params.get("include_transpiration", True))
    include_runoff = bool(params.get("include_runoff", True))
    include_infiltration = bool(params.get("include_infiltration", False))
    show_sun = bool(params.get("show_sun", True))

    fig, ax = plt.subplots(figsize=(8.8, 5.2))
    ax.set_title(plot_title)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")

    # --- Ground / water (collection) ---
    # Water body
    water = plt.Rectangle((0.8, 0.6), 3.2, 1.1, fill=False, linewidth=2)
    ax.add_patch(water)
    ax.text(2.4, 0.25, lab_coll, ha="center", fontsize=11)

    # Land (simple hill)
    land_x = [4.0, 6.0, 7.4, 9.2]
    land_y = [0.6, 1.6, 1.0, 0.6]
    ax.plot(land_x, land_y, linewidth=2)
    ax.plot([4.0, 9.2], [0.6, 0.6], linewidth=2)

    # Tree (symbolic)
    if include_transpiration:
        ax.plot([6.3, 6.3], [1.1, 2.1], linewidth=2)  # trunk
        ax.plot([6.1, 6.3, 6.5], [2.1, 2.6, 2.1], linewidth=2)  # canopy triangle-ish
        ax.text(6.3, 2.95, lab_trans, ha="center", fontsize=10)

    # --- Cloud (condensation) ---
    # Simple cloud: three circles and an outline box-ish using circles
    cloud_centers = [(6.3, 4.5), (7.0, 4.7), (7.7, 4.5)]
    for (cx, cy) in cloud_centers:
        ax.add_patch(plt.Circle((cx, cy), 0.45, fill=False, linewidth=2))
    ax.plot([5.9, 8.1], [4.15, 4.15], linewidth=2)
    ax.text(7.0, 5.35, lab_cond, ha="center", fontsize=11)

    # --- Sun ---
    if show_sun:
        ax.add_patch(plt.Circle((1.4, 4.8), 0.45, fill=False, linewidth=2))
        # rays
        for dx, dy in [(0.0, 0.9), (0.0, -0.9), (0.9, 0.0), (-0.9, 0.0), (0.65, 0.65), (-0.65, 0.65), (0.65, -0.65), (-0.65, -0.65)]:
            ax.plot([1.4, 1.4 + dx], [4.8, 4.8 + dy], linewidth=1.4)

    # --- Evaporation arrow (water up) ---
    ax.annotate(
        "", xy=(2.4, 3.9), xytext=(2.4, 1.75),
        arrowprops=dict(arrowstyle="->", linewidth=2)
    )
    ax.text(2.4, 3.95, lab_evap, ha="center", fontsize=11)

    # --- Condensation move (up/right to cloud) ---
    ax.annotate(
        "", xy=(6.2, 4.3), xytext=(2.8, 4.1),
        arrowprops=dict(arrowstyle="->", linewidth=2)
    )

    # --- Precipitation arrow (cloud down) ---
    ax.annotate(
        "", xy=(7.0, 2.0), xytext=(7.0, 4.1),
        arrowprops=dict(arrowstyle="->", linewidth=2)
    )
    ax.text(7.0, 1.75, lab_prec, ha="center", fontsize=11)

    # Raindrops (symbolic)
    for x in [6.6, 7.0, 7.4]:
        ax.plot([x, x], [3.1, 2.7], linewidth=1.6)
        ax.plot([x, x], [2.6, 2.2], linewidth=1.6)

    # --- Runoff arrow (land back to water) ---
    if include_runoff:
        ax.annotate(
            "", xy=(4.2, 1.05), xytext=(7.9, 1.05),
            arrowprops=dict(arrowstyle="->", linewidth=2)
        )
        ax.text(6.0, 1.25, lab_runoff, ha="center", fontsize=10)

    # --- Infiltration arrow (down into ground) ---
    if include_infiltration:
        ax.annotate(
            "", xy=(6.9, 0.2), xytext=(6.9, 1.1),
            arrowprops=dict(arrowstyle="->", linewidth=2)
        )
        ax.text(6.9, 0.05, lab_infil, ha="center", fontsize=10)

    return _fig_to_png_bytes(fig)


def render_food_web(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic Food Web diagram.
    Required params:
      - nodes: list of nodes, each:
          {"id":"grass", "label":"Grass"}
        or {"id":"rabbit","label":"Rabbit"}
      - links: list of directed links (energy flow), each:
          {"from":"grass", "to":"rabbit"}
        Meaning: grass -> rabbit (rabbit eats grass), energy flows to rabbit.
    Optional params:
      - title
      - layout: "auto" (default) or "levels"
      - levels: dict mapping node_id -> level int (only used if layout="levels")
      - show_legend (bool) default False
    """
    nodes = params["nodes"]
    links = params["links"]
    plot_title = params.get("title") or title or "Food Web"
    layout = (params.get("layout") or "auto").strip().lower()
    show_legend = bool(params.get("show_legend", False))

    if not (isinstance(nodes, list) and nodes):
        raise ValueError("nodes must be a non-empty list.")
    if not (isinstance(links, list) and links):
        raise ValueError("links must be a non-empty list.")

    # Build node maps
    node_ids = []
    labels = {}
    for n in nodes:
        if not isinstance(n, dict):
            raise ValueError("Each node must be a dict with id/label.")
        nid = str(n.get("id", "")).strip()
        lab = str(n.get("label", nid)).strip()
        if not nid:
            raise ValueError("Each node requires a non-empty 'id'.")
        node_ids.append(nid)
        labels[nid] = lab

    node_set = set(node_ids)

    # Validate links
    for e in links:
        if not isinstance(e, dict):
            raise ValueError("Each link must be a dict {'from':..., 'to':...}.")
        a = str(e.get("from", "")).strip()
        b = str(e.get("to", "")).strip()
        if not a or not b:
            raise ValueError("Each link must have 'from' and 'to'.")
        if a not in node_set or b not in node_set:
            raise ValueError(f"Link references unknown node: {a}->{b}")

    # Layout positions
    positions: Dict[str, Tuple[float, float]] = {}

    if layout == "levels":
        # User provides levels mapping
        lvl_map = params.get("levels") or {}
        if not isinstance(lvl_map, dict) or not lvl_map:
            raise ValueError("layout='levels' requires a 'levels' dict mapping node_id -> level.")
        # group by level
        by_lvl: Dict[int, list] = {}
        for nid in node_ids:
            if nid not in lvl_map:
                raise ValueError(f"Node '{nid}' missing from levels mapping.")
            lv = int(lvl_map[nid])
            by_lvl.setdefault(lv, []).append(nid)

        lvls = sorted(by_lvl.keys())
        # vertical spacing by level
        y_top, y_bot = 0.85, 0.15
        if len(lvls) == 1:
            y_positions = {lvls[0]: 0.5}
        else:
            step = (y_top - y_bot) / (len(lvls) - 1)
            y_positions = {lv: (y_top - i * step) for i, lv in enumerate(lvls)}

        # spread nodes across x for each level
        for lv in lvls:
            nids = by_lvl[lv]
            k = len(nids)
            if k == 1:
                xs = [0.5]
            else:
                xs = [0.15 + j * (0.7 / (k - 1)) for j in range(k)]
            for nid, x in zip(nids, xs):
                positions[nid] = (x, y_positions[lv])

    else:
        # Auto layout: simple heuristic by indegree/outdegree
        indeg = {nid: 0 for nid in node_ids}
        outdeg = {nid: 0 for nid in node_ids}
        for e in links:
            outdeg[e["from"]] += 1
            indeg[e["to"]] += 1

        producers = [nid for nid in node_ids if indeg[nid] == 0]
        apex = [nid for nid in node_ids if outdeg[nid] == 0]
        middle = [nid for nid in node_ids if nid not in producers and nid not in apex]

        # y levels: producers bottom, middle mid, apex top
        layers = [(producers, 0.2), (middle, 0.5), (apex, 0.8)]
        for group, y in layers:
            if not group:
                continue
            k = len(group)
            if k == 1:
                xs = [0.5]
            else:
                xs = [0.15 + j * (0.7 / (k - 1)) for j in range(k)]
            for nid, x in zip(group, xs):
                positions[nid] = (x, y)

        # Any node not placed (edge case) -> center
        for nid in node_ids:
            positions.setdefault(nid, (0.5, 0.5))

    # Render
    fig, ax = plt.subplots(figsize=(8.6, 5.4))
    ax.set_title(plot_title)
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    # Draw nodes as rounded boxes
    for nid in node_ids:
        x, y = positions[nid]
        lab = labels[nid]
        ax.text(
            x, y, lab,
            ha="center", va="center", fontsize=11,
            bbox=dict(boxstyle="round,pad=0.35", linewidth=1.8)
        )

    # Draw arrows
    for e in links:
        a = str(e["from"]).strip()
        b = str(e["to"]).strip()
        x1, y1 = positions[a]
        x2, y2 = positions[b]
        ax.annotate(
            "",
            xy=(x2, y2), xytext=(x1, y1),
            arrowprops=dict(arrowstyle="->", linewidth=1.6)
        )

    if show_legend:
        ax.text(0.02, 0.02, "Arrows show energy flow: food → consumer", transform=ax.transAxes, fontsize=10)

    return _fig_to_png_bytes(fig)


def render_circuit_series_parallel(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic circuit diagram: series or parallel.
    Required params:
      - mode: "series" or "parallel"
      - components: list of components in the load section.
          Each component can be:
            {"type":"lamp", "label":"L1"}
            {"type":"resistor", "label":"R1"}
          For series: components are in one line.
          For parallel: components are branches (each component becomes one branch).
    Optional params:
      - title
      - show_switch (bool) default True
      - switch_closed (bool) default True
      - show_labels (bool) default True
      - battery_label (str) default "Battery"
      - voltage_label (str) optional, e.g. "9V"
      - annotate_current (bool) default False  # draws "I" arrow on main line
    """
    mode = (params["mode"] or "").strip().lower()
    comps = params["components"]
    plot_title = params.get("title") or title or "Circuit diagram"

    show_switch = bool(params.get("show_switch", True))
    switch_closed = bool(params.get("switch_closed", True))
    show_labels = bool(params.get("show_labels", True))
    battery_label = params.get("battery_label") or "Battery"
    voltage_label = (params.get("voltage_label") or "").strip()
    annotate_current = bool(params.get("annotate_current", False))

    if mode not in {"series", "parallel"}:
        raise ValueError("mode must be 'series' or 'parallel'.")
    if not isinstance(comps, list) or not comps:
        raise ValueError("components must be a non-empty list.")

    # Normalize components
    norm = []
    for c in comps:
        if not isinstance(c, dict):
            raise ValueError("Each component must be a dict like {'type':'lamp','label':'L1'}.")
        ctype = str(c.get("type", "")).strip().lower()
        lab = str(c.get("label", "")).strip()
        if ctype not in {"lamp", "resistor"}:
            raise ValueError("component type must be 'lamp' or 'resistor'.")
        norm.append((ctype, lab))

    fig, ax = plt.subplots(figsize=(9.0, 4.8 if mode == "parallel" else 4.2))
    ax.set_title(plot_title)
    ax.axis("off")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)

    # Helpers
    def line(x1, y1, x2, y2, lw=2.0):
        ax.plot([x1, x2], [y1, y2], linewidth=lw)

    def text(x, y, s, **kw):
        ax.text(x, y, s, fontsize=10, **kw)

    def draw_battery(x, y, h=1.2, gap=0.25):
        # two vertical plates: long (positive) and short (negative)
        # wires connect left and right
        # plate positions
        x1 = x
        x2 = x + gap
        # long plate
        line(x1, y - h/2, x1, y + h/2, lw=2.2)
        # short plate
        line(x2, y - (h*0.32), x2, y + (h*0.32), lw=2.2)

    def draw_switch(x, y, length=1.2):
        # Simple open/closed switch on a horizontal line
        # left contact
        line(x, y, x + 0.35, y)
        # right contact
        line(x + length - 0.35, y, x + length, y)
        if switch_closed:
            line(x + 0.35, y, x + length - 0.35, y)
        else:
            # open: angled arm not touching
            line(x + 0.35, y, x + length - 0.45, y + 0.25)

    def draw_resistor(x1, y, x2):
        # Rectangle resistor body
        w = x2 - x1
        body_w = min(1.2, w * 0.55)
        left = x1 + (w - body_w) / 2
        right = left + body_w
        # leads
        line(x1, y, left, y)
        line(right, y, x2, y)
        rect = plt.Rectangle((left, y - 0.22), body_w, 0.44, fill=False, linewidth=2.0)
        ax.add_patch(rect)

    def draw_lamp(x1, y, x2):
        # Circle lamp with X
        w = x2 - x1
        r = min(0.35, w * 0.18)
        cx = (x1 + x2) / 2
        # leads
        line(x1, y, cx - r, y)
        line(cx + r, y, x2, y)
        circ = plt.Circle((cx, y), r, fill=False, linewidth=2.0)
        ax.add_patch(circ)
        line(cx - r*0.7, y - r*0.7, cx + r*0.7, y + r*0.7, lw=1.6)
        line(cx - r*0.7, y + r*0.7, cx + r*0.7, y - r*0.7, lw=1.6)

    def draw_component(ctype, label, x1, y, x2):
        if ctype == "resistor":
            draw_resistor(x1, y, x2)
        else:
            draw_lamp(x1, y, x2)
        if show_labels and label:
            text((x1 + x2) / 2, y + 0.55, label, ha="center", va="bottom")

    # Core geometry
    y_main = 3.0
    x_left = 1.0
    x_right = 9.0

    # Left vertical return wire (loop)
    # We'll draw a rectangular-ish loop with battery on left side.
    # Bottom wire
    y_bottom = 1.2
    y_top = 4.8

    # Left side: battery placed mid on left vertical
    line(x_left, y_bottom, x_left, y_top)
    draw_battery(x_left + 0.05, y_main, h=1.4, gap=0.28)

    if show_labels:
        text(x_left - 0.2, y_main + 1.0, battery_label, ha="left", va="bottom")
        if voltage_label:
            text(x_left - 0.2, y_main + 0.7, voltage_label, ha="left", va="bottom")

    # Top wire to the right (with optional switch)
    # from left to switch start
    x_switch_start = 2.3
    x_switch_end = 3.6
    line(x_left, y_top, x_switch_start, y_top)

    if show_switch:
        draw_switch(x_switch_start, y_top, length=(x_switch_end - x_switch_start))
        if show_labels:
            text((x_switch_start + x_switch_end)/2, y_top + 0.45, "Switch", ha="center")
        line(x_switch_end, y_top, x_right, y_top)
    else:
        line(x_switch_start, y_top, x_right, y_top)

    # Right vertical wire down
    line(x_right, y_top, x_right, y_bottom)

    # Bottom wire back to left
    line(x_right, y_bottom, x_left, y_bottom)

    # Now place the "load section" between right side and left side on the midline,
    # but connected into the loop: easiest is to draw a branch from top wire down to midline, then back up.
    # We'll create a "load box" on the right half for clarity.
    x_load_left = 4.2
    x_load_right = 8.2

    # Connect from top wire down to load midline
    line(x_load_left, y_top, x_load_left, y_main)
    # Connect from load midline back up to top wire near right
    line(x_load_right, y_main, x_load_right, y_top)

    # Draw current arrow optionally on the top wire
    if annotate_current:
        ax.annotate("", xy=(6.0, y_top + 0.02), xytext=(5.0, y_top + 0.02),
                    arrowprops=dict(arrowstyle="->", linewidth=2.0))
        text(6.05, y_top + 0.08, "I", ha="left", va="bottom")

    # Load drawing
    if mode == "series":
        # One line of components on y_main between x_load_left and x_load_right
        line(x_load_left, y_main, x_load_right, y_main)

        # Allocate segments for each component
        n = len(norm)
        span = x_load_right - x_load_left
        seg = span / n
        for i, (ctype, lab) in enumerate(norm):
            sx1 = x_load_left + i * seg
            sx2 = x_load_left + (i + 1) * seg
            draw_component(ctype, lab, sx1, y_main, sx2)

    else:
        # Parallel: branches between two vertical rails at x_load_left and x_load_right
        # Draw rails
        line(x_load_left, y_main - 1.6, x_load_left, y_main + 1.6)
        line(x_load_right, y_main - 1.6, x_load_right, y_main + 1.6)

        n = len(norm)
        # distribute branch y positions
        if n == 1:
            ys = [y_main]
        else:
            y1 = y_main + 1.2
            y2 = y_main - 1.2
            step = (y1 - y2) / (n - 1)
            ys = [y1 - i * step for i in range(n)]

        for (ctype, lab), yb in zip(norm, ys):
            # branch wires
            line(x_load_left, yb, x_load_right, yb)
            # component centered in branch
            draw_component(ctype, lab, x_load_left + 0.6, yb, x_load_right - 0.6)

    return _fig_to_png_bytes(fig)


def render_slope_triangle(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic slope triangle on coordinate plane.
    Required params:
      - p1: {"x": x1, "y": y1}
      - p2: {"x": x2, "y": y2}
    Optional:
      - title
      - show_grid (bool) default True
      - label_rise (str) default "rise"
      - label_run (str) default "run"
      - x_min, x_max, y_min, y_max (auto if not provided)
    """
    p1 = params["p1"]
    p2 = params["p2"]

    x1 = float(p1["x"]); y1 = float(p1["y"])
    x2 = float(p2["x"]); y2 = float(p2["y"])

    plot_title = params.get("title") or title or "Slope triangle"
    show_grid = bool(params.get("show_grid", True))
    label_rise = params.get("label_rise") or "rise"
    label_run = params.get("label_run") or "run"

    # Bounds
    xs = [x1, x2]; ys = [y1, y2]
    x_min = float(params.get("x_min", min(xs) - 2))
    x_max = float(params.get("x_max", max(xs) + 2))
    y_min = float(params.get("y_min", min(ys) - 2))
    y_max = float(params.get("y_max", max(ys) + 2))

    fig, ax = plt.subplots(figsize=(6.0, 5.0))
    ax.set_title(plot_title)
    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min, y_max)
    ax.set_xlabel("x")
    ax.set_ylabel("y")

    if show_grid:
        ax.grid(True, linestyle="--", linewidth=0.6)

    # Axes through origin if visible
    if x_min <= 0 <= x_max:
        ax.axvline(0, linewidth=1.2)
    if y_min <= 0 <= y_max:
        ax.axhline(0, linewidth=1.2)

    # Line segment
    ax.plot([x1, x2], [y1, y2], linewidth=2.0, marker="o")

    # Right triangle: from p1 to (x2,y1) to p2
    ax.plot([x1, x2], [y1, y1], linewidth=2.0)
    ax.plot([x2, x2], [y1, y2], linewidth=2.0)

    # Labels run and rise
    mid_run_x = (x1 + x2) / 2
    ax.text(mid_run_x, y1 + 0.3, label_run, ha="center", fontsize=10)

    mid_rise_y = (y1 + y2) / 2
    ax.text(x2 + 0.3, mid_rise_y, label_rise, va="center", fontsize=10)

    return _fig_to_png_bytes(fig)


def render_gradient_rise_run(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Deterministic rise/run arrows between two points.
    Required params:
      - p1: {"x": x1, "y": y1}
      - p2: {"x": x2, "y": y2}
    Optional:
      - title
      - show_grid (bool) default True
      - show_values (bool) default True
      - x_min, x_max, y_min, y_max (auto if not provided)
    """
    p1 = params["p1"]
    p2 = params["p2"]

    x1 = float(p1["x"]); y1 = float(p1["y"])
    x2 = float(p2["x"]); y2 = float(p2["y"])

    plot_title = params.get("title") or title or "Gradient: rise/run"
    show_grid = bool(params.get("show_grid", True))
    show_values = bool(params.get("show_values", True))

    dx = x2 - x1
    dy = y2 - y1

    # Bounds
    xs = [x1, x2]; ys = [y1, y2]
    x_min = float(params.get("x_min", min(xs) - 2))
    x_max = float(params.get("x_max", max(xs) + 2))
    y_min = float(params.get("y_min", min(ys) - 2))
    y_max = float(params.get("y_max", max(ys) + 2))

    fig, ax = plt.subplots(figsize=(6.0, 5.0))
    ax.set_title(plot_title)
    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min, y_max)
    ax.set_xlabel("x")
    ax.set_ylabel("y")

    if show_grid:
        ax.grid(True, linestyle="--", linewidth=0.6)

    if x_min <= 0 <= x_max:
        ax.axvline(0, linewidth=1.2)
    if y_min <= 0 <= y_max:
        ax.axhline(0, linewidth=1.2)

    # Points
    ax.scatter([x1, x2], [y1, y2], s=50)
    ax.text(x1, y1, "  P1", va="bottom")
    ax.text(x2, y2, "  P2", va="bottom")

    # Run arrow (horizontal)
    ax.annotate("", xy=(x2, y1), xytext=(x1, y1), arrowprops=dict(arrowstyle="->", linewidth=2.0))
    # Rise arrow (vertical)
    ax.annotate("", xy=(x2, y2), xytext=(x2, y1), arrowprops=dict(arrowstyle="->", linewidth=2.0))

    # Optional values
    if show_values:
        ax.text((x1 + x2)/2, y1 + 0.3, f"run = {dx:g}", ha="center", fontsize=10)
        ax.text(x2 + 0.3, (y1 + y2)/2, f"rise = {dy:g}", va="center", fontsize=10)
        if dx != 0:
            m = dy / dx
            ax.text(0.02, 0.98, f"gradient = {m:g}", transform=ax.transAxes, ha="left", va="top", fontsize=10)

    return _fig_to_png_bytes(fig)


def render_bar_chart(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    categories = params["categories"]
    values = params["values"]
    plot_title = params.get("title") or title or "Bar chart"
    x_label = params.get("x_label") or ""
    y_label = params.get("y_label") or ""
    units = params.get("units") or ""
    show_grid = bool(params.get("show_grid", False))
    rotate_x = float(params.get("rotate_x", 0))

    if len(categories) != len(values):
        raise ValueError("categories and values must be same length.")

    fig, ax = plt.subplots(figsize=(6.4, 4.0))
    ax.bar(range(len(categories)), values)
    ax.set_xticks(range(len(categories)))
    ax.set_xticklabels(categories, rotation=rotate_x, ha="right" if rotate_x else "center")
    ax.set_title(plot_title)
    ax.set_xlabel(x_label)
    ax.set_ylabel((y_label + (f" ({units})" if units else "")).strip())
    if show_grid:
        ax.grid(True, axis="y", linestyle="--", linewidth=0.6)
    return _fig_to_png_bytes(fig)


def render_line_graph(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    x = params["x"]
    y = params["y"]
    plot_title = params.get("title") or title or "Line graph"
    x_label = params.get("x_label") or ""
    y_label = params.get("y_label") or ""
    units_x = params.get("units_x") or ""
    units_y = params.get("units_y") or ""
    show_grid = bool(params.get("show_grid", True))
    marker = bool(params.get("marker", False))

    if len(x) != len(y):
        raise ValueError("x and y must be same length.")

    fig, ax = plt.subplots(figsize=(6.4, 4.0))
    ax.plot(x, y, marker="o" if marker else None)
    ax.set_title(plot_title)
    ax.set_xlabel((x_label + (f" ({units_x})" if units_x else "")).strip())
    ax.set_ylabel((y_label + (f" ({units_y})" if units_y else "")).strip())
    if show_grid:
        ax.grid(True, linestyle="--", linewidth=0.6)
    return _fig_to_png_bytes(fig)


def render_scatter_plot_blank(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    """
    Blank scatter plot template (axes + grid only).
    Level-aware scaling via defaults_by_level (already applied), with guards here too.
    """
    plot_title = (params.get("title") or title or "Scatter plot").strip()
    x_label = (params.get("x_label") or "").strip()
    y_label = (params.get("y_label") or "").strip()
    show_grid = bool(params.get("show_grid", True))

    # Defaults should be applied already, but guard anyway
    try:
        x_min = float(params.get("x_min", 0))
        x_max = float(params.get("x_max", 10))
        y_min = float(params.get("y_min", 0))
        y_max = float(params.get("y_max", 10))
        x_step = float(params.get("x_step", 1))
        y_step = float(params.get("y_step", 1))
    except Exception:
        x_min, x_max, y_min, y_max, x_step, y_step = 0, 10, 0, 10, 1, 1

    if x_max <= x_min:
        x_max = x_min + 10
    if y_max <= y_min:
        y_max = y_min + 10
    if x_step <= 0:
        x_step = 1
    if y_step <= 0:
        y_step = 1

    fig = plt.figure(figsize=(7.2, 5.2), dpi=220)
    ax = fig.add_axes([0.12, 0.14, 0.83, 0.78])

    ax.set_title(plot_title)
    if x_label:
        ax.set_xlabel(x_label)
    if y_label:
        ax.set_ylabel(y_label)

    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min, y_max)

    # Ticks
    xticks = []
    t = x_min
    while t <= x_max + 1e-9:
        xticks.append(t)
        t += x_step
    yticks = []
    t = y_min
    while t <= y_max + 1e-9:
        yticks.append(t)
        t += y_step

    ax.set_xticks(xticks)
    ax.set_yticks(yticks)

    if show_grid:
        ax.grid(True, linestyle="--", alpha=0.35)

    return _fig_to_png_bytes(fig)


def render_scatter_plot(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    x = params["x"]
    y = params["y"]
    plot_title = params.get("title") or title or "Scatter plot"
    x_label = params.get("x_label") or ""
    y_label = params.get("y_label") or ""
    units_x = params.get("units_x") or ""
    units_y = params.get("units_y") or ""
    show_grid = bool(params.get("show_grid", True))
    best_fit_line = bool(params.get("best_fit_line", False))

    if len(x) != len(y):
        raise ValueError("x and y must be same length.")

    fig, ax = plt.subplots(figsize=(6.4, 4.0))
    ax.scatter(x, y)
    ax.set_title(plot_title)
    ax.set_xlabel((x_label + (f" ({units_x})" if units_x else "")).strip())
    ax.set_ylabel((y_label + (f" ({units_y})" if units_y else "")).strip())

    if best_fit_line and len(x) >= 2:
        n = len(x)
        x_mean = sum(x) / n
        y_mean = sum(y) / n
        denom = sum((xi - x_mean) ** 2 for xi in x) or 1e-9
        m = sum((x[i] - x_mean) * (y[i] - y_mean) for i in range(n)) / denom
        b = y_mean - m * x_mean
        x_min, x_max = min(x), max(x)
        ax.plot([x_min, x_max], [m * x_min + b, m * x_max + b])

    if show_grid:
        ax.grid(True, linestyle="--", linewidth=0.6)

    return _fig_to_png_bytes(fig)


def render_pie_chart(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    labels = params["labels"]
    values = params["values"]
    plot_title = params.get("title") or title or "Pie chart"
    show_percent = bool(params.get("show_percent", True))

    if len(labels) != len(values):
        raise ValueError("labels and values must be same length.")
    total = sum(values)
    if total <= 0:
        raise ValueError("values must sum to > 0.")

    fig, ax = plt.subplots(figsize=(6.0, 4.0))
    autopct = "%1.0f%%" if show_percent else None
    ax.pie(values, labels=labels, autopct=autopct)
    ax.set_title(plot_title)
    return _fig_to_png_bytes(fig)


def render_number_line(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    vmin = float(params["min"])
    vmax = float(params["max"])
    ticks = int(params.get("ticks", 6))
    mark_points = params.get("mark_points", []) or []
    plot_title = params.get("title") or title or "Number line"

    if vmax <= vmin:
        raise ValueError("max must be greater than min.")
    ticks = max(2, ticks)

    fig, ax = plt.subplots(figsize=(8.0, 2.0))
    ax.set_title(plot_title)
    ax.hlines(0, vmin, vmax, linewidth=2)
    ax.set_ylim(-1, 1)
    ax.set_yticks([])

    step = (vmax - vmin) / (ticks - 1)
    tick_vals = [vmin + i * step for i in range(ticks)]
    ax.set_xticks(tick_vals)
    ax.set_xlim(vmin, vmax)

    for p in mark_points:
        if isinstance(p, dict):
            x = float(p.get("x"))
            lab = str(p.get("label", "")).strip()
        else:
            x = float(p)
            lab = ""
        ax.vlines(x, -0.15, 0.15, linewidth=2)
        if lab:
            ax.text(x, 0.25, lab, ha="center", va="bottom", fontsize=10)

    ax.spines["left"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["top"].set_visible(False)

    return _fig_to_png_bytes(fig)


def render_free_body_diagram(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    forces = params["forces"]
    plot_title = params.get("title") or title or "Free body diagram"
    show_axes = bool(params.get("show_axes", True))

    if not isinstance(forces, list) or not forces:
        raise ValueError("forces must be a non-empty list.")

    fig, ax = plt.subplots(figsize=(4.5, 4.5))
    ax.set_title(plot_title)
    ax.plot([0], [0], marker="o")

    if show_axes:
        ax.axhline(0, linewidth=1)
        ax.axvline(0, linewidth=1)

    max_mag = max(float(f.get("magnitude", 1)) for f in forces if isinstance(f, dict)) or 1.0
    scale = 1.0 / max_mag

    for f in forces:
        if not isinstance(f, dict):
            raise ValueError("Each force must be a dict like {'label':..., 'angle_deg':..., 'magnitude':...}")
        lab = str(f.get("label", "F")).strip()
        ang = float(f.get("angle_deg", 0.0))
        mag = float(f.get("magnitude", 1.0))
        r = mag * scale * 1.8
        dx = r * math.cos(math.radians(ang))
        dy = r * math.sin(math.radians(ang))
        ax.arrow(0, 0, dx, dy, length_includes_head=True, head_width=0.08, head_length=0.12)
        ax.text(dx * 1.08, dy * 1.08, lab, ha="center", va="center", fontsize=10)

    ax.set_xlim(-2.2, 2.2)
    ax.set_ylim(-2.2, 2.2)
    ax.set_aspect("equal", "box")
    ax.set_xticks([])
    ax.set_yticks([])

    return _fig_to_png_bytes(fig)



def render_simple_circuit(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    plot_title = params.get("title") or title or "Simple circuit"
    closed_switch = bool(params.get("closed_switch", True))

    fig, ax = plt.subplots(figsize=(6.0, 2.8))
    ax.set_title(plot_title)
    ax.axis("off")

    x0, x1 = 0.1, 0.9
    y_top, y_bot = 0.85, 0.15

    # --- positions ---
    y_cell_mid = 0.67
    gap_half = 0.10
    gap_low = y_cell_mid - gap_half
    gap_high = y_cell_mid + gap_half

    bulb_x = 0.5
    bulb_r = 0.05

    sx = 0.75
    sw_half = 0.06

    # --- Wires with gaps ---
    # Left rail split around cell (prevents wire-through-plates)
    ax.plot([x0, x0], [y_bot, gap_low], linewidth=2)
    ax.plot([x0, x0], [gap_high, y_top], linewidth=2)

    # Right rail full
    ax.plot([x1, x1], [y_bot, y_top], linewidth=2)

    # Top wire split around switch
    ax.plot([x0, sx - sw_half], [y_top, y_top], linewidth=2)
    ax.plot([sx + sw_half, x1], [y_top, y_top], linewidth=2)

    # Bottom wire split around bulb
    ax.plot([x0, bulb_x - bulb_r], [y_bot, y_bot], linewidth=2)
    ax.plot([bulb_x + bulb_r, x1], [y_bot, y_bot], linewidth=2)

    # --- Cell symbol (on the gap) ---
    bx = x0
    ax.plot([bx - 0.02, bx + 0.02], [y_cell_mid - 0.05, y_cell_mid - 0.05], linewidth=2)  # short plate
    ax.plot([bx - 0.03, bx + 0.03], [y_cell_mid + 0.05, y_cell_mid + 0.05], linewidth=2)  # long plate
    ax.text(bx, y_cell_mid + 0.10, "Cell", ha="center", fontsize=9)

    # --- Bulb symbol (wire already gapped) ---
    ax.add_patch(plt.Circle((bulb_x, y_bot), bulb_r, fill=False, linewidth=2))
    ax.text(bulb_x, y_bot - 0.11, "Bulb", ha="center", fontsize=9)

    # --- Switch symbol (wire already gapped) ---
    if closed_switch:
        ax.plot([sx - sw_half, sx + sw_half], [y_top, y_top], linewidth=2)
        ax.text(sx, y_top + 0.07, "Switch (closed)", ha="center", fontsize=8)
    else:
        ax.plot([sx - sw_half, sx - 0.01], [y_top, y_top], linewidth=2)
        ax.plot([sx + 0.01, sx + sw_half], [y_top, y_top], linewidth=2)
        ax.plot([sx - 0.01, sx + 0.03], [y_top, y_top + 0.06], linewidth=2)
        ax.text(sx, y_top + 0.07, "Switch (open)", ha="center", fontsize=8)

    return _fig_to_png_bytes(fig)


def render_series_circuit(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    loads = params["loads"]
    plot_title = params.get("title") or title or "Series circuit"
    if not isinstance(loads, list) or not loads:
        raise ValueError("loads must be a non-empty list.")

    fig, ax = plt.subplots(figsize=(7.0, 2.8))
    ax.set_title(plot_title)
    ax.axis("off")

    x0, x1 = 0.1, 0.9
    y_top, y_bot = 0.8, 0.2

    # Cell gap on left rail
    y_cell_mid = 0.60
    gap_half = 0.10
    gap_low = y_cell_mid - gap_half
    gap_high = y_cell_mid + gap_half

    ax.plot([x0, x0], [y_bot, gap_low], linewidth=2)
    ax.plot([x0, x0], [gap_high, y_top], linewidth=2)
    ax.plot([x1, x1], [y_bot, y_top], linewidth=2)

    # Top wire full
    ax.plot([x0, x1], [y_top, y_top], linewidth=2)

    # Bottom wire with gaps around load rectangles
    n = len(loads)
    xs = [x0 + (i + 1) * (x1 - x0) / (n + 1) for i in range(n)]
    comp_w = 0.10
    comp_h = 0.06

    prev = x0
    for cx in xs:
        ax.plot([prev, cx - comp_w / 2], [y_bot, y_bot], linewidth=2)
        prev = cx + comp_w / 2
    ax.plot([prev, x1], [y_bot, y_bot], linewidth=2)

    # Cell plates on the gap
    bx = x0
    ax.plot([bx - 0.02, bx + 0.02], [y_cell_mid - 0.05, y_cell_mid - 0.05], linewidth=2)
    ax.plot([bx - 0.03, bx + 0.03], [y_cell_mid + 0.05, y_cell_mid + 0.05], linewidth=2)
    ax.text(bx, y_cell_mid + 0.10, "Cell", ha="center", fontsize=9)

    for i, comp in enumerate(loads):
        lab = str(comp).strip().title()
        cx = xs[i]
        ax.add_patch(plt.Rectangle((cx - comp_w / 2, y_bot - comp_h / 2), comp_w, comp_h, fill=False, linewidth=2))
        ax.text(cx, y_bot - 0.12, lab, ha="center", fontsize=9)

    return _fig_to_png_bytes(fig)


def render_parallel_circuit(params: Dict[str, Any], level: str, subject: str, title: Optional[str]) -> bytes:
    branches = params["branches"]
    plot_title = params.get("title") or title or "Parallel circuit"
    if not isinstance(branches, list) or not branches:
        raise ValueError("branches must be a non-empty list (each branch is a list of loads).")

    fig, ax = plt.subplots(figsize=(7.0, 3.5))
    ax.set_title(plot_title)
    ax.axis("off")

    x_left, x_right = 0.2, 0.8
    y_top, y_bot = 0.85, 0.15

    # Cell gap on left rail
    y_cell_mid = 0.60
    gap_half = 0.10
    gap_low = y_cell_mid - gap_half
    gap_high = y_cell_mid + gap_half

    ax.plot([x_left, x_left], [y_bot, gap_low], linewidth=2)
    ax.plot([x_left, x_left], [gap_high, y_top], linewidth=2)
    ax.plot([x_right, x_right], [y_bot, y_top], linewidth=2)

    # Cell symbol
    bx = x_left
    ax.plot([bx - 0.02, bx + 0.02], [y_cell_mid - 0.05, y_cell_mid - 0.05], linewidth=2)
    ax.plot([bx - 0.03, bx + 0.03], [y_cell_mid + 0.05, y_cell_mid + 0.05], linewidth=2)
    ax.text(bx, y_cell_mid + 0.10, "Cell", ha="center", fontsize=9)

    # Branches with gaps around loads
    m = len(branches)
    ys = [y_top - (i + 1) * (y_top - y_bot) / (m + 1) for i in range(m)]
    load_w = 0.26
    load_h = 0.08
    cx = (x_left + x_right) / 2

    for bi, branch in enumerate(branches):
        y = ys[bi]
        ax.plot([x_left, cx - load_w / 2], [y, y], linewidth=2)
        ax.plot([cx + load_w / 2, x_right], [y, y], linewidth=2)

        if not isinstance(branch, list) or not branch:
            branch = ["Load"]
        label = " + ".join(str(x).strip().title() for x in branch)

        ax.add_patch(plt.Rectangle((cx - load_w / 2, y - load_h / 2), load_w, load_h, fill=False, linewidth=2))
        ax.text(cx, y, label, ha="center", va="center", fontsize=9)

    return _fig_to_png_bytes(fig)

def list_inventory(group_by: str = "subject") -> Dict[str, Any]:
    """
    Returns the Master Diagram Archetype Inventory as a structured dict.
    group_by:
      - "subject" (default): subjects -> archetypes
      - "level": levels -> archetypes
      - "none": flat list
    Each archetype entry includes:
      - archetype_id
      - title
      - subjects
      - levels
      - deterministic
      - implemented (renderer callable if deterministic)
      - fallback (if non-deterministic)
      - required_params / optional_params
    """
    flat = []

    for aid, meta in ARCHETYPES.items():
        deterministic = bool(meta.get("deterministic", False))
        renderer_name = meta.get("renderer")
        implemented = False
        if deterministic:
            implemented = callable(globals().get(renderer_name)) if renderer_name else False

        flat.append({
            "archetype_id": aid,
            "title": meta.get("title", ""),
            "subjects": meta.get("subjects", []) or [],
            "levels": meta.get("levels", []) or [],
            "deterministic": deterministic,
            "implemented": implemented,
            "fallback": meta.get("fallback", None) if not deterministic else None,
            "required_params": meta.get("required_params", []) or [],
            "optional_params": meta.get("optional_params", []) or [],
        })

    # Sort consistently for UI
    flat.sort(key=lambda x: (",".join(x["subjects"]), x["title"], x["archetype_id"]))

    if group_by == "none":
        return {"group_by": "none", "items": flat}

    if group_by == "level":
        out: Dict[str, Any] = {"group_by": "level", "groups": {}}
        for item in flat:
            lvls = item.get("levels") or ["(unspecified)"]
            for lvl in lvls:
                out["groups"].setdefault(lvl, []).append(item)
        # sort within each level
        for lvl in out["groups"]:
            out["groups"][lvl].sort(key=lambda x: (x["deterministic"] is False, x["implemented"] is False, x["title"]))
        return out

    # default: group_by == "subject"
    out2: Dict[str, Any] = {"group_by": "subject", "groups": {}}
    for item in flat:
        subs = item.get("subjects") or ["(unspecified)"]
        for sub in subs:
            out2["groups"].setdefault(sub, []).append(item)

    for sub in out2["groups"]:
        out2["groups"][sub].sort(key=lambda x: (x["deterministic"] is False, x["implemented"] is False, x["title"]))
    return out2


def inventory_summary() -> Dict[str, Any]:
    """
    Small summary numbers for dashboards/QA.
    """
    total = len(ARCHETYPES)
    deterministic_total = 0
    deterministic_implemented = 0
    fallback_total = 0

    for _, meta in ARCHETYPES.items():
        if meta.get("deterministic", False):
            deterministic_total += 1
            rname = meta.get("renderer")
            if rname and callable(globals().get(rname)):
                deterministic_implemented += 1
        else:
            fallback_total += 1

    return {
        "total_archetypes": total,
        "deterministic_total": deterministic_total,
        "deterministic_implemented": deterministic_implemented,
        "deterministic_not_implemented_yet": deterministic_total - deterministic_implemented,
        "fallback_total": fallback_total,
    }

# ADD THIS SECTION to diagram_library.py (Step 10)
# -----------------------------------------------
# Goal: Internal QA for the Master Inventory so you catch mistakes early.
# This is safe to call at app start (or in a CI test).

def validate_inventory(strict: bool = False) -> Dict[str, Any]:
    """
    Validates the ARCHETYPES registry.
    strict=False (default):
      - warnings are returned but do not raise
    strict=True:
      - raises ValueError if any errors found
    Checks:
      - archetype_id integrity
      - required keys exist (title, subjects, levels, deterministic)
      - deterministic archetypes have a renderer string
      - deterministic renderer functions exist OR are at least named (warn if missing)
      - required_params / optional_params are lists
      - synonyms is a list (if present)
    """
    errors = []
    warnings = []

    if not isinstance(ARCHETYPES, dict) or not ARCHETYPES:
        errors.append("ARCHETYPES must be a non-empty dict.")
        return _finalize_validation(errors, warnings, strict)

    for aid, meta in ARCHETYPES.items():
        if not isinstance(aid, str) or not aid.strip():
            errors.append(f"Invalid archetype_id key: {repr(aid)}")
            continue
        if not isinstance(meta, dict):
            errors.append(f"Archetype '{aid}' meta must be a dict.")
            continue

        # Required fields
        for key in ["title", "subjects", "levels", "deterministic"]:
            if key not in meta:
                errors.append(f"Archetype '{aid}' missing required key '{key}'.")

        # Type checks
        if "subjects" in meta and not isinstance(meta.get("subjects"), list):
            errors.append(f"Archetype '{aid}' subjects must be a list.")
        if "levels" in meta and not isinstance(meta.get("levels"), list):
            errors.append(f"Archetype '{aid}' levels must be a list.")
        if "required_params" in meta and not isinstance(meta.get("required_params"), list):
            errors.append(f"Archetype '{aid}' required_params must be a list.")
        if "optional_params" in meta and not isinstance(meta.get("optional_params"), list):
            errors.append(f"Archetype '{aid}' optional_params must be a list.")
        if "synonyms" in meta and not isinstance(meta.get("synonyms"), list):
            errors.append(f"Archetype '{aid}' synonyms must be a list.")

        deterministic = bool(meta.get("deterministic", False))

        if deterministic:
            renderer = meta.get("renderer", None)
            if not renderer or not isinstance(renderer, str):
                errors.append(f"Archetype '{aid}' is deterministic but missing a valid 'renderer' string.")
            else:
                # Renderer exists?
                fn = globals().get(renderer)
                if not callable(fn):
                    warnings.append(f"Archetype '{aid}' renderer '{renderer}' is not implemented (will fallback).")
        else:
            # fallback should exist (warn if not)
            if not meta.get("fallback"):
                warnings.append(f"Archetype '{aid}' is non-deterministic but has no 'fallback' specified.")

        # Title sanity
        title = meta.get("title", "")
        if isinstance(title, str) and not title.strip():
            warnings.append(f"Archetype '{aid}' has an empty title.")

        # Synonym sanity (weak check for “too generic”)
        syns = meta.get("synonyms") or []
        if syns:
            for s in syns:
                if isinstance(s, str) and s.strip() in {"diagram", "graph", "chart"}:
                    warnings.append(f"Archetype '{aid}' contains very generic synonym '{s}' (may cause misrouting).")

    return _finalize_validation(errors, warnings, strict)


def _finalize_validation(errors: list, warnings: list, strict: bool) -> Dict[str, Any]:
    report = {
        "ok": (len(errors) == 0),
        "error_count": len(errors),
        "warning_count": len(warnings),
        "errors": errors,
        "warnings": warnings,
    }
    if strict and errors:
        raise ValueError("Inventory validation failed: " + "; ".join(errors[:5]))
    return report


# ADD THIS SECTION to diagram_library.py (Step 20 - Final)
# -------------------------------------------------------
# Goal: quick deterministic smoke tests you can run on Hugging Face after deploy.
# It generates one example for each new archetype and reports pass/fail.
#
# How to use:
#   report = run_smoke_tests()
#   print(report)

def run_smoke_tests() -> Dict[str, Any]:
    """
    Generates one deterministic PNG per archetype and reports pass/fail.
    Does NOT touch paid fallback generation.
    Returns:
      {
        "ok": bool,
        "passed": [...],
        "failed": [{"archetype":..., "error":...}, ...]
      }
    """
    tests = [
        ("histogram", {
            "bins": [0, 10, 20, 30],
            "frequencies": [3, 5, 2],
            "title": "Histogram test"
        }),
        ("box_and_whisker", {
            "min": 2, "q1": 5, "median": 7, "q3": 10, "max": 14,
            "show_outliers": True, "outliers": [1, 16],
            "title": "Box plot test"
        }),
        ("coordinate_plane_points", {
            "points": [{"x": 1, "y": 2, "label": "A"}, {"x": -2, "y": 1, "label": "B"}],
            "title": "Points test"
        }),
        ("frequency_table", {
            "rows": [{"label": "Cats", "freq": 5}, {"label": "Dogs", "freq": 3}],
            "show_tally": True,
            "title": "Frequency table test"
        }),
        ("two_way_table", {
            "row_labels": ["Male", "Female"],
            "col_labels": ["Yes", "No"],
            "values_matrix": [[12, 8], [9, 11]],
            "title": "Two-way table test"
        }),
        ("probability_tree", {
            "stages": [
                {"name": "Coin", "branches": [{"label": "Heads", "p": "1/2"}, {"label": "Tails", "p": "1/2"}]},
                {"name": "Spinner", "branches": [{"label": "Red", "p": "1/3"}, {"label": "Blue", "p": "2/3"}]},
            ],
            "show_joint": True,
            "title": "Probability tree test"
        }),
        ("venn_2", {
            "a_only": 7, "b_only": 5, "both": 3,
            "label_a": "A", "label_b": "B",
            "title": "Venn test"
        }),
        ("stem_and_leaf", {
            "data": [12, 14, 15, 18, 21, 22, 27, 27, 33],
            "title": "Stem-and-leaf test"
        }),
        ("slope_triangle", {
            "p1": {"x": 1, "y": 2},
            "p2": {"x": 5, "y": 6},
            "title": "Slope triangle test"
        }),
        ("gradient_rise_run", {
            "p1": {"x": 1, "y": 2},
            "p2": {"x": 5, "y": 6},
            "title": "Rise/Run test"
        }),
        ("water_cycle", {
            "title": "Water cycle test"
        }),
        ("food_web", {
            "nodes": [
                {"id": "grass", "label": "Grass"},
                {"id": "rabbit", "label": "Rabbit"},
                {"id": "fox", "label": "Fox"},
            ],
            "links": [
                {"from": "grass", "to": "rabbit"},
                {"from": "rabbit", "to": "fox"},
            ],
            "title": "Food web test",
            "show_legend": True
        }),
        ("circuit_series_parallel", {
            "mode": "series",
            "components": [{"type": "lamp", "label": "L1"}, {"type": "lamp", "label": "L2"}],
            "title": "Circuit test",
            "show_switch": True,
            "switch_closed": True
        }),
    ]

    passed = []
    failed = []

    for archetype_id, params in tests:
        try:
            req = {
                "prompt": f"smoke test {archetype_id}",
                "subject": (ARCHETYPES.get(archetype_id, {}).get("subjects") or ["math"])[0],
                "level": "J",
                "archetype_hint": archetype_id,
                "params": params,
            }
            res = generate_diagram(req, user_ctx=None)
            if res.get("status") != "ok":
                raise RuntimeError(f"generate_diagram returned status={res.get('status')} reason={res.get('reason')}")
            png_bytes = res.get("bytes")
            if not png_bytes or not isinstance(png_bytes, (bytes, bytearray)):
                raise RuntimeError("No png_bytes returned.")
            if len(png_bytes) < 2000:
                raise RuntimeError(f"png_bytes too small ({len(png_bytes)} bytes)")

            passed.append(archetype_id)
        except Exception as e:
            failed.append({"archetype": archetype_id, "error": str(e)})

    return {
        "ok": (len(failed) == 0),
        "passed": passed,
        "failed": failed,
    }




# === CIRCUIT ARCHETYPE OVERRIDES (AUTO-PATCH) ===
# This block updates circuit archetypes without depending on the exact formatting above.
# It runs after ARCHETYPES is defined.

try:
    ARCHETYPES.update({
        "simple_circuit": {
            "title": "Simple circuit (cell, switch, lamp)",
            "subjects": ["science", "physics"],
            "levels": ["J", "S"],
            "deterministic": True,
            "renderer": "render_simple_circuit",
            "required_params": [],
            "optional_params": ["closed_switch", "title"],
            "defaults_by_level": {
                "J": {"closed_switch": True, "title": "Simple circuit"},
                "S": {"closed_switch": False, "title": "Simple circuit (open switch)"},
            },
        },
        "series_circuit": {
            "title": "Series circuit",
            "subjects": ["science", "physics"],
            "levels": ["J", "S"],
            "deterministic": True,
            "renderer": "render_series_circuit",
            "required_params": ["loads"],
            "optional_params": ["title"],
            "defaults_by_level": {
                "J": {"loads": ["Lamp"], "title": "Series circuit"},
                "S": {"loads": ["Lamp", "Lamp"], "title": "Series circuit (two lamps)"},
            },
        },
        "parallel_circuit": {
            "title": "Parallel circuit",
            "subjects": ["science", "physics"],
            "levels": ["J", "S"],
            "deterministic": True,
            "renderer": "render_parallel_circuit",
            "required_params": ["branches"],
            "optional_params": ["title"],
            "defaults_by_level": {
                "J": {"branches": [["Lamp"], ["Lamp"]], "title": "Parallel circuit"},
                "S": {"branches": [["Lamp"], ["Lamp"], ["Lamp"]], "title": "Parallel circuit (three branches)"},
            },
        },
        "series_parallel_circuit": {
            "title": "Series-parallel circuit",
            "subjects": ["science", "physics"],
            "levels": ["J", "S"],
            "deterministic": True,
            "renderer": "render_parallel_circuit",
            "required_params": ["branches"],
            "optional_params": ["title"],
            "defaults_by_level": {
                "J": {"branches": [["Lamp", "Lamp"], ["Lamp"]], "title": "Series-parallel circuit"},
                "S": {"branches": [["Lamp", "Lamp"], ["Lamp", "Lamp"]], "title": "Series-parallel circuit (two branches)"},
            },
        },
    })
except Exception as e:
    print("⚠️ CIRCUIT ARCHETYPE OVERRIDES failed:", repr(e))


# ---------------------------------------------------------------------------
### AUTO_PATCH_CIRCUITS_V28
# Deterministic defaults for circuit archetypes + deterministic series-parallel renderer.
# This prevents fallback to image_gen when params are not provided in [[VISUAL ...]].
# ---------------------------------------------------------------------------


def render_series_parallel_circuit(params: dict, level: str, subject: str, title: str | None):
    series_loads = params.get("series_loads") or ["Bulb"]
    branches = params.get("branches") or [["Bulb"], ["Bulb"]]
    plot_title = params.get("title") or title or "Series-parallel circuit"

    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.set_title(plot_title)
    ax.axis("off")

    xL, xR = 0.2, 0.8
    yTop, yBot = 0.85, 0.15

    # Cell gap on left rail
    y_cell_mid = 0.60
    gap_half = 0.10
    gap_low = y_cell_mid - gap_half
    gap_high = y_cell_mid + gap_half

    ax.plot([xL, xL], [yBot, gap_low], linewidth=2)
    ax.plot([xL, xL], [gap_high, yTop], linewidth=2)
    ax.plot([xR, xR], [yBot, yTop], linewidth=2)

    # Cell symbol
    bx = xL
    ax.plot([bx - 0.02, bx + 0.02], [y_cell_mid - 0.05, y_cell_mid - 0.05], linewidth=2)
    ax.plot([bx - 0.03, bx + 0.03], [y_cell_mid + 0.05, y_cell_mid + 0.05], linewidth=2)
    ax.text(bx, y_cell_mid + 0.10, "Cell", ha="center", fontsize=9)

    # Top wire
    ax.plot([xL, xR], [yTop, yTop], linewidth=2)

    # Bottom series wire with gaps around series loads
    n = max(1, len(series_loads))
    xs = [xL + (i + 1) * (xR - xL) / (n + 1) for i in range(n)]
    comp_w = 0.12
    comp_h = 0.06

    prev = xL
    for cx in xs:
        ax.plot([prev, cx - comp_w / 2], [yBot, yBot], linewidth=2)
        prev = cx + comp_w / 2
    ax.plot([prev, xR], [yBot, yBot], linewidth=2)

    for i, comp in enumerate(series_loads):
        lab = str(comp).strip().title()
        cx = xs[i]
        ax.add_patch(plt.Rectangle((cx - comp_w / 2, yBot - comp_h / 2), comp_w, comp_h, fill=False, linewidth=2))
        ax.text(cx, yBot - 0.12, lab, ha="center", fontsize=9)

    # Parallel section baseline
    ySplit = 0.62
    ax.plot([xL, xR], [ySplit, ySplit], linewidth=2)

    # Parallel branches with gaps around loads
    m = max(1, len(branches))
    ys = [yTop - (i + 1) * (yTop - ySplit) / (m + 1) for i in range(m)]
    load_w = 0.30
    load_h = 0.08
    cx = (xL + xR) / 2

    for bi, branch in enumerate(branches):
        y = ys[bi]
        ax.plot([xL, cx - load_w / 2], [y, y], linewidth=2)
        ax.plot([cx + load_w / 2, xR], [y, y], linewidth=2)

        if not isinstance(branch, list) or not branch:
            branch = ["Load"]
        label = " + ".join(str(x).strip().title() for x in branch)

        ax.add_patch(plt.Rectangle((cx - load_w / 2, y - load_h / 2), load_w, load_h, fill=False, linewidth=2))
        ax.text(cx, y, label, ha="center", va="center", fontsize=9)

    return _fig_to_png_bytes(fig)

def _apply_circuit_defaults_v28():
    if not isinstance(globals().get("ARCHETYPES"), dict):
        return

    # series_circuit: ensure defaults for loads
    sc = ARCHETYPES.get("series_circuit")
    if isinstance(sc, dict):
        sc.setdefault("defaults_by_level", {})
        for lvl in ("J", "S", "U"):
            sc["defaults_by_level"].setdefault(lvl, {})
            sc["defaults_by_level"][lvl].setdefault("loads", ["Bulb", "Bulb"])
        sc["required_params"] = ["loads"]

    # parallel_circuit: ensure defaults for branches
    pc = ARCHETYPES.get("parallel_circuit")
    if isinstance(pc, dict):
        pc.setdefault("defaults_by_level", {})
        for lvl in ("J", "S", "U"):
            pc["defaults_by_level"].setdefault(lvl, {})
            pc["defaults_by_level"][lvl].setdefault("branches", [["Bulb"], ["Bulb"]])
        pc["required_params"] = ["branches"]

    # series_parallel_circuit: force deterministic renderer + defaults
    sp = ARCHETYPES.get("series_parallel_circuit")
    if isinstance(sp, dict):
        sp.setdefault("defaults_by_level", {})
        for lvl in ("J", "S", "U"):
            sp["defaults_by_level"].setdefault(lvl, {})
            sp["defaults_by_level"][lvl].setdefault("series_loads", ["Bulb"])
            sp["defaults_by_level"][lvl].setdefault("branches", [["Bulb"], ["Bulb"]])
        sp["deterministic"] = True
        sp["renderer"] = "render_series_parallel_circuit"
        sp["required_params"] = []  # renderer is safe with defaults


_apply_circuit_defaults_v28()
# ---------------------------------------------------------------------------
