# ======================================================================================
# credits.py
# ======================================================================================
# Module: Credits, Wallet, Ledger & Usage Cost Control Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module manages the credit-based billing control layer of EduDraft Studio.
#
# It is responsible for wallet initialization, balance retrieval, direct credit spending,
# ledger recording, and lightweight credit estimation for visual-generation workflows.
#
# It provides the operational bridge between user actions, stored credit balances,
# and paid system capabilities such as image generation.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Wallet Initialization
#    - Ensures each authenticated user has an associated wallet record
#    - Retrieves the current wallet balance through Supabase RPC flows
#
# 2. Balance Lookup
#    - Fetches the user's current available credit balance
#    - Returns consistent balance/error tuples for UI and system use
#
# 3. Credit Spending
#    - Performs direct debit operations against the user's credit wallet
#    - Prevents spending when available balance is insufficient
#    - Updates both:
#        • credit_wallet
#        • credit_ledger
#
# 4. Ledger Recording
#    - Writes each credit spend as a ledger event
#    - Stores reason and metadata for billing transparency and auditability
#
# 5. Visual Credit Estimation
#    - Parses Markdown for [[VISUAL ...]] placeholders
#    - Estimates how many paid image-type visuals will require credits
#    - Calculates required credits before execution
#
# 6. UI-Friendly Status Reporting
#    - Returns balance summary text for in-app display
#    - Supports wallet visibility and user-facing billing clarity
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - Supabase → wallet, ledger, and RPC interactions
# - auth.py → authenticated session validation
# - config.py → feature flags and credit-cost settings
# - regex parsing → VISUAL placeholder inspection
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Billing safety first
# - No spending without sufficient balance
# - Simple, explainable wallet and ledger flows
# - Clear separation between balance lookup and spend execution
# - Stable credit estimation rules for predictable behaviour
#
# --------------------------------------------------------------------------------------
# DATABASE / RPC CONTRACT
# --------------------------------------------------------------------------------------
# RPC functions:
#   - credit_ensure_wallet
#   - credit_balance
#
# Tables:
#   - credit_wallet
#   - credit_ledger
#
# Expected wallet model:
#   - one wallet row per user
#   - running balance stored in credit_wallet
#   - immutable spend history stored in credit_ledger
#
# --------------------------------------------------------------------------------------
# SYSTEM ROLE
# --------------------------------------------------------------------------------------
# This module is the commercial enforcement layer for paid capabilities.
#
# In practice:
#   - diagram_library.py may require credits before image generation fallback
#   - app.py / UI flows can display current wallet balance
#   - billing-sensitive actions can use this module to enforce paid access safely
#
# It therefore sits at the intersection of:
#   - monetisation
#   - usage control
#   - auditability
#   - feature gating
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - Current estimation logic only charges for VISUAL placeholders with:
#     kind="image" | "photo" | "picture"
# - This can later be extended to diagram fallbacks when image generation is triggered
# - Any changes here affect both user billing behaviour and platform monetisation logic
#
# ======================================================================================

import re
from typing import Dict, Tuple, Any
from datetime import datetime, timezone

from config import supabase, ENABLE_IMAGE_GEN, IMAGE_CREDITS_PER_IMAGE
from auth import _require_session

_VISUAL_RE = re.compile(r"\[\[VISUAL\s+([^\]]+)\]\]", re.IGNORECASE)
_KV_RE = re.compile(r'(\w+)\s*=\s*"([^"]*)"')


def _parse_visual_kv(payload: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for m in _KV_RE.finditer(payload or ""):
        out[m.group(1).strip().lower()] = m.group(2)
    return out


def _rpc_balance(res) -> float:
    """
    Our RPCs return: [{"balance": <number>}]
    """
    data = getattr(res, "data", None)
    if isinstance(data, list) and data and isinstance(data[0], dict) and "balance" in data[0]:
        try:
            return float(data[0]["balance"] or 0)
        except Exception:
            return 0.0
    return 0.0


def ensure_wallet(sess) -> Tuple[float, str]:
    """
    Ensures a wallet row exists and returns (balance, err_msg).
    """
    _a, _r, _uid, err = _require_session(sess)
    if err:
        return 0.0, err

    try:
        res = supabase.rpc("credit_ensure_wallet", {}).execute()
        return _rpc_balance(res), ""
    except Exception as e:
        return 0.0, f"❌ Credits ensure failed: {type(e).__name__}: {e}"


def get_balance(sess) -> Tuple[float, str]:
    """
    Returns (balance, err_msg).
    """
    _a, _r, _uid, err = _require_session(sess)
    if err:
        return 0.0, err

    try:
        res = supabase.rpc("credit_balance", {}).execute()
        return _rpc_balance(res), ""
    except Exception as e:
        return 0.0, f"❌ Credits lookup failed: {type(e).__name__}: {e}"


def spend_credits(user_id: str, amount: float, reason: str, meta: dict | None = None, sb=None) -> tuple[bool, str, float]:
    """
    Spend credits directly via wallet + ledger tables.
    Returns (ok, msg, balance).

    If an authenticated Supabase client is supplied via sb, use it so RLS runs
    under the active user's session. Otherwise fall back to the global client.
    """
    if not user_id:
        return False, "Missing user_id", 0.0

    spend_amount = float(abs(amount))
    db = sb or supabase

    try:
        res = db.table("credit_wallet").select("balance").eq("user_id", user_id).execute()
        data = getattr(res, "data", None)
        if isinstance(data, list) and data and isinstance(data[0], dict):
            balance = float(data[0].get("balance") or 0)
        else:
            balance = 0.0

        if balance < spend_amount:
            return False, "Insufficient credits", balance

        new_balance = balance - spend_amount

        db.table("credit_ledger").insert({
            "user_id": user_id,
            "amount": -spend_amount,
            "reason": reason,
            "meta": meta or {},
        }).execute()

        db.table("credit_wallet").update({
            "balance": new_balance,
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }).eq("user_id", user_id).execute()

        return True, "OK", new_balance
    except Exception as e:
        return False, f"❌ Credit spend failed: {type(e).__name__}: {e}", 0.0


def credits_needed_for_markdown(md_text: str) -> Tuple[int, float]:
    """
    V2.6 rule (simple + stable):
    - Only placeholders with kind="image" require image-gen credits.
    Later we can expand to kind="diagram" when it falls back to image-gen.
    """
    md = md_text or ""
    count = 0

    for m in _VISUAL_RE.finditer(md):
        payload = (m.group(1) or "").strip()
        kv = _parse_visual_kv(payload)
        kind = (kv.get("kind", "") or "").strip().lower()

        if kind in {"image", "photo", "picture"}:
            count += 1

    required = float(count) * float(IMAGE_CREDITS_PER_IMAGE)
    return count, required


def credits_status_text(sess) -> str:
    bal, err = ensure_wallet(sess)
    if err:
        return err
    return f"Balance: {bal:.2f}"
