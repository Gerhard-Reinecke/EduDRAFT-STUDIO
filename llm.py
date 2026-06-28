# ======================================================================================
# llm.py
# ======================================================================================
# Module: LLM Orchestration, Prompt Engineering & Output Control Engine
#
# System: EduDraft Studio (Marike App)
# Version: 4.4.10
#
# --------------------------------------------------------------------------------------
# PURPOSE
# --------------------------------------------------------------------------------------
# This module is the central intelligence layer of EduDraft Studio.
#
# It orchestrates all interactions with the OpenAI API, constructs structured prompts,
# enforces strict output contracts, and transforms raw model responses into reliable,
# application-ready document components.
#
# It ensures that all AI-generated content is deterministic, structured, and safe for
# downstream rendering, storage, and editing workflows.
#
# --------------------------------------------------------------------------------------
# CORE RESPONSIBILITIES
# --------------------------------------------------------------------------------------
# 1. Prompt Engineering & Context Construction
#    - Builds structured teacher context from UI inputs
#    - Enforces global teaching assistant behaviour via SYSTEM_PROMPT
#    - Injects curriculum-aware metadata (country, level, subject, etc.)
#    - Applies strict VISUAL placeholder rules for diagrams and graphics
#
# 2. LLM Invocation Layer
#    - Interfaces with OpenAI models (responses API or chat completions fallback)
#    - Supports both generation mode and edit mode workflows
#    - Appends a strict output contract to enforce structured responses
#
# 3. Output Structure Enforcement
#    - Ensures all responses follow:
#         DOCUMENT_MARKDOWN
#         PPT_OUTLINE
#         ANSWER_KEY
#    - Detects malformed outputs and auto-recovers missing sections
#    - Prevents empty editor states by generating safe fallback content
#
# 4. Response Parsing & Section Extraction
#    - Splits model output into document, slides, and memo components
#    - Supports multiple formatting styles (label-based, heading-based, fallback parsing)
#    - Ensures robustness against model inconsistencies
#
# 5. Document Composition
#    - Combines document content and answer key when required
#    - Maintains consistent formatting for downstream export and rendering
#
# 6. Audio Transcription Support
#    - Converts uploaded audio into text using OpenAI transcription models
#    - Enables voice-driven document generation workflows
#
# --------------------------------------------------------------------------------------
# DEPENDENCIES
# --------------------------------------------------------------------------------------
# - OpenAI client (via config.py) → LLM interaction
# - Supabase → user/session linkage (indirect via auth)
# - auth.py → session validation
# - re → parsing and fallback extraction logic
#
# --------------------------------------------------------------------------------------
# DESIGN PRINCIPLES
# --------------------------------------------------------------------------------------
# - Deterministic output (strict contract enforcement)
# - Defensive programming against LLM unpredictability
# - Zero empty-state tolerance (always return usable content)
# - Teacher-first prompt design (clear, structured, contextual)
# - Model-agnostic compatibility (supports multiple OpenAI SDK patterns)
#
# --------------------------------------------------------------------------------------
# CRITICAL RULES ENFORCED
# --------------------------------------------------------------------------------------
# - Output MUST contain:
#     DOCUMENT_MARKDOWN
#     PPT_OUTLINE
#     ANSWER_KEY
#
# - Visuals MUST be represented as [[VISUAL ...]] placeholders only
# - No free-form AI narration outside structured sections
# - No "As an AI..." or conversational filler
#
# --------------------------------------------------------------------------------------
# NOTES
# --------------------------------------------------------------------------------------
# - This module is tightly coupled with all ingestion pipelines (PDF, DOCX, PPTX)
# - Any change here affects the entire generation system
# - Fallback logic ensures app stability even if the LLM misbehaves
# - Visual placeholder rules are critical for downstream diagram rendering
#
# ======================================================================================


import re
from datetime import datetime, timezone

from config import client, supabase, DAILY_LIMIT_GENERATE, DAILY_LIMIT_TRANSCRIBE
from auth import _require_session

# =============================
# OPENAI + DOCUMENT LOGIC
# =============================
SYSTEM_PROMPT = r"""
You are a global teaching assistant for educators worldwide across ALL subjects.
ROLE
- Draft tests, quizzes, worksheets, investigations, lesson notes, worked examples, marking keys/memos/memorandums/answer sheets, rubrics, and PowerPoint slide outlines across ALL subjects.
- Support School (Primary/Secondary) and University/Tertiary levels.
- Use the educator's selected Country and (if provided) State/Province/Region to match curriculum conventions and spelling (e.g., maths/math, marks/points, year/grade).
- No "As an AI..." lines. No waffle.
SUBJECT RULE
- The teacher’s selected Subject/Course controls the content domain (e.g., Maths, English, Biology, Science, Computer Science, Engineering, History, Geography, Languages, etc.).
- If the teacher’s instruction clearly indicates a different subject than the selected Subject/Course, follow the teacher’s instruction.
EDUCATION LEVEL RULE
- If Education level is School (Primary / Secondary): use school-appropriate tone, scaffolding, and assessment style. Follow national/state curriculum conventions as closely as possible.
- If Education level is University / Tertiary: assume lecturer autonomy, higher rigor, formal exposition, and university-style assessments (tutorial sheets, problem sets, exams, lecture outlines).
MATH NOTATION AND ASSESSMENT VALIDITY
- If mathematical expressions appear in the content, format math using LaTeX:
  Inline: \( ... \)
  Display: \[ ... \]
- Convert spoken math into correct notation.
- Never leave maths expressions, values, variables, options, probabilities, units, table values, or answers blank unless the blank is intentionally for the student to complete.
- For generated tests/exams/quizzes, all numerical data must be internally consistent.
- Do not generate impossible totals, contradictory probabilities, mismatched answer keys, or memo corrections such as "assume typo", "check", "wait", "recompute", or "depends".
- Do not reveal reasoning, self-correction, uncertainty, drafting notes, or internal checks in DOCUMENT_MARKDOWN or ANSWER_KEY.
- If a generated question would be inconsistent, silently choose corrected values before writing the final question and memo.
- Only add an assumption line if the teacher’s instruction is genuinely missing essential information; do not add assumptions to fix your own generated content.
OUTPUT RULE
Return EXACTLY this structure:
1) DOCUMENT_MARKDOWN:
<clean classroom-ready content in Markdown>
2) PPT_OUTLINE:
Slide 1: <title>
- bullets...
Speaker notes: ...
Slide 2: ...
3) ANSWER_KEY (only if requested):
<solutions / memo in Markdown>
Do not include anything else.
VISUAL PLACEHOLDER ENFORCEMENT:
If any visual is requested, you MUST insert a [[VISUAL ...]] placeholder at the correct position.
Never draw or describe the visual in normal text outside the placeholder.
""".strip()

EDIT_GUARDRAILS = r"""
You are editing an EXISTING teacher document.
STRICT RULES:
- Do NOT rewrite the whole document.
- Only apply the requested changes.
- Keep structure, numbering, marks/points, and formatting unchanged unless the teacher explicitly asks to change them.
- If the teacher request is ambiguous, make the smallest reasonable change and add 1 short note at the end: "Assumption:".
Return EXACTLY the same output structure as usual:
1) DOCUMENT_MARKDOWN:
...
2) PPT_OUTLINE:
...
3) ANSWER_KEY (only if requested):
...
""".strip()


def safe_err(msg: str, e: Exception) -> str:
    return f"{msg}\n\nERROR:\n{type(e).__name__}: {e}"


def transcribe_audio(audio_path: str) -> str:
    with open(audio_path, "rb") as f:
        tr = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return tr.text.strip()


def build_user_request(
    instruction_text: str,
    education_level: str,
    country: str,
    state_province: str,
    year_level: str,
    course: str,
    output_type: str,
    include_memo: bool,
    university_name: str = "",
    faculty: str = "",
    module_code: str = "",
    curriculum_stream: str = "",
) -> str:
    extras = []
    extras.append(f"Education level: {education_level}")
    extras.append(f"Country: {country}")

    sp = (state_province or "").strip()
    if sp:
        extras.append(f"State/Province/Region: {sp}")

    if education_level == "School (Primary / Secondary)":
        extras.append(f"Year/Grade: {year_level}")
        if (course or "").strip() and course not in {"Not specified", "Choose Subject"}:
            extras.append(f"Course: {course}")

    cs = (curriculum_stream or "").strip()

    # School: curriculum stream is optional, and must NOT control university metadata.
    if education_level == "School (Primary / Secondary)":
        if cs:
            extras.append(f"Curriculum stream/syllabus: {cs}")
    
    # University: show university fields, independently of curriculum stream.
    else:
        uni = (university_name or "").strip()
        if uni:
            extras.append(f"University/Institution: {uni}")
    
        fac = (faculty or "").strip()
        if fac:
            extras.append(f"Faculty/Discipline: {fac}")
    
        mod = (module_code or "").strip()
        if mod:
            extras.append(f"Course/Module: {mod}")
    
        # (Optional) allow curriculum stream to exist in university too, if you ever want it.
        if cs:
            extras.append(f"Curriculum stream/syllabus: {cs}")


    extras.append(f"Output type: {output_type}")
    extras.append(f"Include answer key/memo: {'Yes' if include_memo else 'No'}")

    extras_block = "\n".join(f"- {x}" for x in extras)

    VISUAL_RULES = """
VISUALS RULE (STRICT, MUST FOLLOW):
If the teacher requests ANY visual (image, diagram, chart, graph, table, map, timeline, picture, illustration, “draw/show a…”, “include a…”, etc.):
A) OUTPUT AS PLACEHOLDER ONLY
- Do NOT draw the visual in text.
- Do NOT describe the visual in prose (except inside the placeholder prompt="...").
- Do NOT include URLs.
- Do NOT claim you “attached” or “generated” an image.
B) INSERT PLACEHOLDER AT THE EXACT LOCATION
- Insert the placeholder EXACTLY where the visual should appear in the document.
- If the visual belongs to a question, place it inside that question.
- If the visual belongs to slides, place it inside the slide where it’s mentioned.
C) PLACEHOLDER FORMAT (ONE SINGLE LINE)
Use EXACTLY this single-line format:
[[VISUAL id="v1" kind="image|diagram|chart|graph|table|map|timeline" where="Q4 or Slide 3" prompt="what to show" notes="style constraints" data="only if chart/table"]]
Rules for fields:
- id: v1, v2, v3… sequential.
- kind: choose the closest kind.
- where: “Q4”, “Q4(b)”, “Section B”, “Slide 3”, etc.
- prompt: clear, teacher-friendly description of what to show.
- notes: style constraints (e.g., “black-and-white, simple labels, classroom-ready”).
- data: ONLY for charts/tables when numbers/categories are provided, e.g. data="A:40,B:35,C:25" or data="Year:2020,2021|Value:12,18"
D) DATA RULES (CRITICAL FOR GRAPHS / STATS)
1) If a question requires INTERPRETING a graph/chart (e.g., “interpret the histogram”, “answer questions from the histogram”, “use the scatter plot to...”), you MUST provide a complete dataset inside the VISUAL placeholder so the chart can be rendered.
- Histogram: include explicit intervals AND frequencies in prompt (preferred) or data field.
  Example prompt: "Histogram of jump lengths. Intervals: 100-109,110-119,120-129,130-139,140-149,150-159. Frequencies: 1,3,6,8,5,2."
- Stem-and-leaf (given): include the raw data list OR the stem|leaf rows.
  Example prompt: "Stem-and-leaf plot for data: 12,14,15,17,19,20,21,21,23,24,26,28..."
- Box plot (given): include five-number summary OR raw data.
- Scatter plot (given): include points as (x,y) pairs.
2) If a question asks students to CONSTRUCT/DRAW the graph themselves (e.g., “draw a histogram”, “complete the stem-and-leaf”, “construct a box plot”), you MUST request a BLANK TEMPLATE visual (axes/grid/template), and you MUST NOT provide the completed data visualization unless the teacher explicitly asked for a worked example.
- Use prompt wording like: "Blank histogram axes template for students to draw bars."
3) If the teacher requests a chart/graph but provides no data AND the task requires interpretation, you MUST invent a small, realistic dataset suitable for the selected year/grade and the topic, and include it in the VISUAL prompt. Do NOT use data="MISSING" for interpretation tasks.
4) Only use data="MISSING" when the teacher explicitly says “use my data” but does not provide it.
Then add ONE line at end: "Assumption: Chart data not provided."
E) MAX ASSUMPTIONS
- If placement/meaning is ambiguous, make the smallest reasonable assumption and add exactly one “Assumption:” line at the end.
""".strip()

    prompt = f"""
Teacher context:
{extras_block}
Teacher instruction (verbatim):
\"\"\"{instruction_text}\"\"\"
{VISUAL_RULES}
Now produce the requested resource following the OUTPUT RULE.
""".strip()

    return prompt


def call_llm(prompt: str, model_name: str, edit_mode: bool = False) -> str:
    """
    Calls the model and enforces the required output structure.
    If the model returns PPT-only (common failure), we wrap it into the required
    DOCUMENT_MARKDOWN / PPT_OUTLINE / ANSWER_KEY format so the app never ends up
    with an empty document editor.
    """
    system_text = SYSTEM_PROMPT if not edit_mode else (SYSTEM_PROMPT + "\n\n" + EDIT_GUARDRAILS)

    # Hard “output contract” appended to user prompt (belt + suspenders)
    contract = (
        "\n\nIMPORTANT OUTPUT CONTRACT (MUST FOLLOW EXACTLY):\n"
        "Return EXACTLY these sections in this order, with these exact labels:\n"
        "DOCUMENT_MARKDOWN:\n"
        "<markdown content>\n\n"
        "PPT_OUTLINE:\n"
        "Slide 1: ...\n\n"
        "ANSWER_KEY:\n"
        "<memo/solutions, or leave blank if not requested>\n\n"
        "Do not add any other text before, between, or after these sections.\n"
    )

    # --- OpenAI SDK compatibility ---
    # Newer SDKs have client.responses.create; older SDKs use client.chat.completions.create
    if hasattr(client, "responses"):
        r = client.responses.create(
            model=model_name,
            input=[
                {"role": "system", "content": system_text},
                {"role": "user", "content": (prompt + contract)},
            ],
        )
        out = (getattr(r, "output_text", "") or "").strip()
    else:
        resp = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_text},
                {"role": "user", "content": (prompt + contract)},
            ],
        )
        out = (resp.choices[0].message.content or "").strip()
    if not out:
        return ""

    # If it already contains the required structure, return as-is
    if "DOCUMENT_MARKDOWN:" in out and "PPT_OUTLINE:" in out:
        return out

    # If we got PPT-ish output but no document section, force-wrap it
    # Extract PPT from "Slide 1:" onward if present
    m = re.search(r"(?is)(^|\n)(Slide\s*1\s*:.*)$", out)
    ppt_part = m.group(2).strip() if m else out

    # Minimal doc fallback (keeps the app functional even if the model misbehaves)
    doc_fallback = (
        "# Draft\n\n"
        "_(Auto-recovered: the model returned slides but no document body.)_\n\n"
        "## Teacher instruction\n"
        "See Transcript/Input above.\n"
    )

    # If the teacher asked for a visual placeholder in the prompt, preserve it in the doc
    vis = re.search(r"(\[\[VISUAL[^\]]+\]\])", prompt, flags=re.I)
    if vis:
        doc_fallback += "\n\n" + vis.group(1).strip() + "\n"

    return (
        "DOCUMENT_MARKDOWN:\n"
        f"{doc_fallback}\n\n"
        "PPT_OUTLINE:\n"
        f"{ppt_part}\n\n"
        "ANSWER_KEY:\n"
    )


def split_sections(llm_text: str):
    """
    Robustly split LLM output into (doc_md, ppt_outline, answer_key).

    Handles:
    - Numbered labels: 1) DOCUMENT_MARKDOWN: ...
    - Unnumbered labels: DOCUMENT_MARKDOWN: ...
    - Headings: ## Document / ## PPT Outline / ## Answer Key
    - Fallback: peel out Slide 1 block
    """
    import re

    text = (llm_text or "").strip()
    if not text:
        return "", "", ""

    # ---------- 1) Scan for LABEL headers using finditer (no lookahead games) ----------
    header_re = re.compile(
        r"(?im)^\s*(?:\d+\s*[\)\.]\s*)?(DOCUMENT_MARKDOWN|PPT_OUTLINE|ANSWER_KEY)\s*:\s*$"
    )

    matches = list(header_re.finditer(text))
    if matches:
        blocks = {}
        for i, m in enumerate(matches):
            label = m.group(1).upper()
            start = m.end()  # content starts after the label line
            end = matches[i + 1].start() if (i + 1 < len(matches)) else len(text)
            blocks[label] = text[start:end].strip()

        return (
            blocks.get("DOCUMENT_MARKDOWN", "").strip(),
            blocks.get("PPT_OUTLINE", "").strip(),
            blocks.get("ANSWER_KEY", "").strip(),
        )

    # ---------- 2) Heading formats ----------
    def grab_heading(pats):
        for pat in pats:
            m = re.search(rf"(?is)(^|\n)\s*#{1,6}\s*{pat}\s*\n(.*?)(?=\n\s*#{1,6}\s|\Z)", text)
            if m:
                return m.group(2).strip()
        return ""

    doc_md = grab_heading([r"(document|document markdown|worksheet|test|lesson|notes)"])
    ppt = grab_heading([r"(ppt|ppt outline|powerpoint|slides|slide outline)"])
    ans = grab_heading([r"(answer key|memo|marking key|solutions)"])

    if doc_md:
        return doc_md, ppt, ans

    # ---------- 3) Fallback: peel out Slide 1 block ----------
    m = re.search(r"(?is)(^|\n)(slide\s*1\s*:\s.*)$", text)
    if m:
        ppt_guess = m.group(2).strip()
        doc_guess = text[:m.start(2)].strip()
        return doc_guess, ppt_guess, ""

    return text, "", ""


def combine_doc_and_memo(doc_md: str, answer_key: str, include_memo: bool) -> str:
    doc_md = (doc_md or "").strip()
    answer_key = (answer_key or "").strip()

    if include_memo and answer_key:
        return doc_md + "\n\n---\n\n# Answer Key / Memo\n\n" + answer_key

    return doc_md


    