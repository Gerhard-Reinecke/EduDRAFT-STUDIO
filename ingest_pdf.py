# ======================================================================================
# ingest_pdf.py
# ======================================================================================
# Module: PDF Ingestion & Draft Reconstruction Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module handles the ingestion, extraction, and conversion of teacher-uploaded
# PDF files into clean, editable educational drafts.
#
# It performs best-effort PDF text extraction using multiple parser fallbacks, rebuilds
# the extracted material into structured Markdown with LaTeX via the LLM pipeline, and
# saves the result into the Drafts system with automatic versioning and preview support.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. PDF Text Extraction
#    - Reads uploaded PDF files from the Gradio workflow
#    - Attempts extraction through multiple fallback libraries
#    - Normalises extracted text for downstream LLM use
#    - Applies a maximum character cap to protect token usage
#
# 2. Upload Analysis
#    - Validates uploaded PDF files and file paths
#    - Reports extraction status, file size, and character count
#    - Enables or disables the generation path based on readability
#
# 3. Draft Reconstruction (LLM Integration)
#    - Converts extracted PDF text into an editable classroom-ready draft
#    - Preserves question numbering, marks, headings, and structure where possible
#    - Supports optional memo / answer key generation
#
# 4. Draft Persistence
#    - Creates a new Draft record in Supabase
#    - Auto-saves Version 1 into draft_versions
#    - Stores contextual metadata such as subject, year level, and curriculum stream
#
# 5. Preview Generation
#    - Builds an HTML preview using MathJax-compatible rendering
#    - Supports immediate review of generated output before further editing
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - pypdf / PyPDF2 / pdfplumber / PyMuPDF → PDF text extraction fallbacks
# - llm.py → Prompt construction and LLM interaction
# - exports.py → HTML preview rendering
# - auth.py → Session validation
# - rate_limit.py → Usage control
# - config.py → Supabase client
# - Gradio → UI interaction layer
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Teacher-first workflow with minimal friction
# - Robust extraction through layered parser fallback strategy
# - Safe token handling through controlled truncation
# - Clean separation between ingestion, generation, preview, and persistence
# - Functional symmetry with DOCX ingestion for a consistent user experience
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - Extraction is best-effort only and depends on the quality of the source PDF
# - Image-only or scanned PDFs may return no readable text
# - OCR is not performed in this module
# - Large extracted outputs are truncated to preserve LLM performance
# - This module mirrors DOCX ingestion behavior where possible for workflow consistency
#
# ======================================================================================



import os
import re
import uuid
import tempfile
from datetime import datetime, timezone

import gradio as gr

from llm import build_user_request, call_llm, split_sections, combine_doc_and_memo
from exports import build_mathjax_html
from auth import _require_session
from rate_limit import check_rate_limit
from config import supabase


# =============================
# PDF INGESTION
# =============================
def extract_text_from_pdf(pdf_path: str, max_chars: int = 120_000) -> str:
    """
    Robust PDF text extraction with multiple fallbacks.
    Returns plain text (best-effort). Truncates to max_chars.
    """
    if not pdf_path or not os.path.exists(pdf_path):
        return ""

    text = ""

    # 1) pypdf
    try:
        from pypdf import PdfReader
        reader = PdfReader(pdf_path)
        parts = [(p.extract_text() or "") for p in reader.pages]
        text = "\n\n".join(parts).strip()
    except Exception:
        pass

    # 2) PyPDF2
    if not text:
        try:
            import PyPDF2
            with open(pdf_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                parts = [(p.extract_text() or "") for p in reader.pages]
            text = "\n\n".join(parts).strip()
        except Exception:
            pass

    # 3) pdfplumber
    if not text:
        try:
            import pdfplumber
            parts = []
            with pdfplumber.open(pdf_path) as pdf:
                for p in pdf.pages:
                    parts.append(p.extract_text() or "")
            text = "\n\n".join(parts).strip()
        except Exception:
            pass

    # 4) PyMuPDF
    if not text:
        try:
            import fitz
            doc = fitz.open(pdf_path)
            parts = [p.get_text("text") or "" for p in doc]
            text = "\n\n".join(parts).strip()
        except Exception:
            pass

    if not text:
        return ""

    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    if len(text) > max_chars:
        text = text[:max_chars] + "\n\n[TRUNCATED]"

    return text


def action_analyze_pdf(pdf_file):
    """
    Called when a PDF is uploaded.
    Returns:
      1) markdown status string
      2) gr.update(interactive=True/False)
    """
    try:
        if not pdf_file:
            return "No PDF uploaded yet.", gr.update(interactive=False)

        if isinstance(pdf_file, dict):
            pdf_path = pdf_file.get("name") or pdf_file.get("path") or ""
            pdf_name = os.path.basename(pdf_path) if pdf_path else "uploaded.pdf"
        else:
            pdf_path = str(pdf_file)
            pdf_name = os.path.basename(pdf_path)

        if not pdf_path or not os.path.exists(pdf_path):
            return "⚠️ PDF uploaded, but file path not found.", gr.update(interactive=False)

        size_kb = os.path.getsize(pdf_path) / 1024.0
        txt = extract_text_from_pdf(pdf_path)
        n = len(txt or "")

        if n == 0:
            msg = (
                f"📄 **PDF detected**\n\n"
                f"- File: `{pdf_name}`\n"
                f"- Size: **{size_kb:.1f} KB**\n"
                f"- Extracted chars: **0** ❌\n\n"
                "Likely scanned or image-only PDF."
            )
            return msg, gr.update(interactive=False)

        msg = (
            f"📄 **PDF detected**\n\n"
            f"- File: `{pdf_name}`\n"
            f"- Size: **{size_kb:.1f} KB**\n"
            f"- Extracted chars: **{n}** ✅"
        )
        return msg, gr.update(interactive=True)

    except Exception as e:
        return f"⚠️ PDF analysis failed: {type(e).__name__}: {e}", gr.update(interactive=False)


def pdf_to_draft_prompt(pdf_text: str, output_type: str, include_memo: bool) -> str:
    memo_line = "Yes" if include_memo else "No"
    return f"""
You are given text extracted from a teacher's PDF. The teacher wants an EDITABLE draft.
Your job:
- Reconstruct the document as a clean classroom-ready editable draft in Markdown with LaTeX.
- Preserve numbering, marks, and structure as best as possible.
- Make minimal assumptions if something is unclear.
Target output type: {output_type}
Include answer key/memo: {memo_line}
PDF_TEXT_START
<<<
{pdf_text}
>>>
PDF_TEXT_END
Now produce the required output following the OUTPUT RULE exactly.
""".strip()


def action_generate_from_pdf(
    supabase_session,
    pdf_file,
    pdf_draft_title: str,
    education_level,
    country, state_province,
    uni_country, university_name, faculty, module_code,
    year_level, course, typed_subject, course_stream, output_type,
    include_memo, model_name
):
    try:
        access_token, refresh_token, user_id, err = _require_session(supabase_session)
        if err:
            return "", "", "", None, f"❌ {err}", "", "PDF upload", "", "", 0

        can_proceed, limit_msg = check_rate_limit(supabase_session, action="generate")
        if not can_proceed:
            return "", "", "", None, limit_msg, "", "PDF upload", "", "", 0

        if not pdf_file:
            return "", "", "", None, f"❌ Please upload a PDF first. {limit_msg}", "", "PDF upload", "", "", 0

        if isinstance(pdf_file, dict):
            pdf_path = pdf_file.get("name") or pdf_file.get("path") or ""
            pdf_name = os.path.basename(pdf_path) if pdf_path else "uploaded.pdf"
        else:
            pdf_path = str(pdf_file)
            pdf_name = os.path.basename(pdf_path)

        if not pdf_path or not os.path.exists(pdf_path):
            return "", "", "", None, f"❌ Uploaded PDF path not found. {limit_msg}", "", "PDF upload", "", "", 0

        pdf_text = extract_text_from_pdf(pdf_path)
        if not (pdf_text or "").strip():
            return "", "", "", None, f"❌ Could not extract text from PDF. {limit_msg}", "", "PDF upload", "", "", 0

        # -----------------------------
        # SUBJECT (GLOBAL, NOT AU-ONLY)
        # -----------------------------
        if education_level == "University / Tertiary":
            eff_country = (uni_country or "").strip() or "Not specified"
            eff_state = ""
            effective_course = (module_code or "").strip() or "Not specified"
        else:
            eff_country = (country or "").strip() or "Not specified"
            eff_state = (state_province or "").strip()

            chosen = (course or "").strip()
            if chosen == "Other (type it)":
                ts = (typed_subject or "").strip()
                if not ts:
                    return "", "", "", None, f"❌ Subject is required. {limit_msg}", "", "PDF upload", "", "", 0
                effective_course = ts
            else:
                effective_course = chosen or "Not specified"

        curriculum_stream = (course_stream or "").strip()

        instruction_text = f"Generate an editable draft from the uploaded PDF: {pdf_name}"

        teacher_context = build_user_request(
            instruction_text=instruction_text,
            education_level=education_level,
            country=eff_country,
            state_province=eff_state,
            year_level=year_level,
            course=effective_course,
            output_type=output_type,
            include_memo=include_memo,
            university_name=university_name,
            faculty=faculty,
            module_code=module_code,
            curriculum_stream=curriculum_stream,
        )

        prompt = teacher_context + "\n\n" + pdf_to_draft_prompt(pdf_text, output_type, include_memo)
        llm_text = call_llm(prompt, model_name, edit_mode=False)

        doc_md, ppt_outline, answer_key = split_sections(llm_text)
        combined_md = combine_doc_and_memo(doc_md, answer_key, include_memo)

        tmpdir = tempfile.mkdtemp()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        html_path = os.path.join(tmpdir, f"Preview_{stamp}.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(build_mathjax_html(combined_md))

        now = datetime.now(timezone.utc).isoformat()
        title = (pdf_draft_title or "").strip() or os.path.splitext(pdf_name)[0]

        draft_id = str(uuid.uuid4())

        draft_row = {
            "id": draft_id,
            "user_id": user_id,
            "title": title,
            "subject": (effective_course or "").strip() or None,
            "curriculum_stream": (curriculum_stream or "").strip() or None,
            "education_level": education_level or None,
            "country": eff_country or None,
            "state_province": eff_state or None,
            "year_level": year_level or None,
            "course": effective_course or None,
            "created_at": now,
            "updated_at": now,
        }
        supabase.table("drafts").insert(draft_row).execute()

        version_row = {
            "draft_id": draft_id,
            "version": 1,
            "doc_md": combined_md or "",
            "ppt_outline": ppt_outline or "",
            "created_at": now,
        }
        supabase.table("draft_versions").insert(version_row).execute()

        status = (
            "✅ PDF converted into a NEW editable draft and auto-saved as Version 1.\n"
            f"Draft name: {title}\n{limit_msg}"
        )

        return (
            f"PDF used: {pdf_name}",
            combined_md,
            ppt_outline,
            html_path,
            status,
            instruction_text,
            "PDF upload",
            f"PDF used: {pdf_name}",
            draft_id,
            1
        )

    except Exception as e:
        from llm import safe_err
        return "", "", "", None, safe_err("PDF → Draft failed.", e), "", "PDF upload", "", "", 0