"""
document_agent.py — Document Agent
Handles doc_task_docx, doc_task_pptx, and doc_task_pdf intents.

Pipeline:
  extractor + style_parser (parallel) → LLM build → validator

- PPTX output is decided by INTENT alone (not source file type)
- PDF is supported as a source-only format (always output docx or pptx)
- template_doc_path is passed to the appropriate builder
"""

import re
import asyncio
import json
import logging
import tempfile
import os
from pathlib import Path
from langchain_core.messages import HumanMessage, SystemMessage

from src.graph.state import AgentState
from src.document.extractor import extract
from src.document.style_parser import extract_styles, extract_pptx_theme
from src.document.builder import build_docx
from src.document.docx_template_builder import build_docx_from_template
from src.document.ppt_builder import build_pptx
from src.document.validator import validate_output
from src.llm.router import get_llm

logger = logging.getLogger(__name__)

# ── Temp directory (cross-platform) ──────────────────────────────────────────

def _get_output_dir(session_id: str) -> str:
    """Return a platform-safe temp directory for this session."""
    base = os.path.join(tempfile.gettempdir(), "docuforge", session_id)
    os.makedirs(base, exist_ok=True)
    return base


# ── LLM System Prompts ────────────────────────────────────────────────────────

DOCX_BUILD_SYSTEM = """You are a document reformatting specialist.

You will receive:
1. Extracted content from a source document (sections, paragraphs, tables)
2. Style metadata from a template document (fonts, margins, headings)
3. User requirements

Your job: produce a structured list of content blocks for building the output .docx.

Return ONLY valid JSON — a list of block objects. No preamble. No markdown fences.

Block types:
- {"type": "heading", "level": 1-6, "text": "..."}
- {"type": "paragraph", "text": "..."}
- {"type": "bullet_list", "items": ["item1", "item2"]}
- {"type": "numbered_list", "items": ["item1", "item2"]}
- {"type": "table", "rows": [["H1","H2"], ["R1C1","R1C2"]]}
- {"type": "page_break"}

Preserve all content from the source. Apply template structure and requirements.
Fix any spelling errors found. Apply all formatting requirements exactly.
"""

PPTX_BUILD_SYSTEM = """You are a presentation design specialist.

You will receive:
1. Extracted content from a source presentation — slides may include "tables" key with structured data
2. Theme data (fonts, colors, font sizes, background) — apply this to ALL slides
3. User requirements

Your job: produce a structured list of slides for the output .pptx.

Return ONLY valid JSON — a list of slide objects. No preamble. No markdown fences.

CRITICAL RULES:
1. Every slide MUST include a "theme" object with font/color data.
2. If a source slide has a "tables" key, you MUST reproduce that table in "table" — NEVER drop table data.
3. Do NOT merge multiple source slides into one — keep each source slide as its own output slide.
4. Do NOT hallucinate or invent new slides. Only use content from the source.

Slide object format:
{
  "layout": "title" | "title_content" | "blank",
  "title": "...",

  // Use "content" for bullet/text slides:
  "content": "string or list of bullet strings",

  // Use "table" for data slides that had a table — MUTUALLY EXCLUSIVE with "content":
  "table": {
    "headers": ["Column1", "Column2", "Column3"],
    "rows": [
      ["Row1Col1", "Row1Col2", "Row1Col3"],
      ["Row2Col1", "Row2Col2", "Row2Col3"]
    ]
  },

  "notes": "optional speaker notes",

  "theme": {
    "fonts": ["Primary Font", "Body Font"],
    "font_sizes": [title_size_pt, body_size_pt],
    "text_colors": ["RRGGBB_title", "RRGGBB_body"],
    "background_color": "RRGGBB or null"
  }
}

Additional rules:
- First slide uses layout "title"
- Section-divider slides (no body) use layout "title"
- Data/table slides use layout "blank" (table is added programmatically)
- Apply theme CONSISTENTLY to ALL slides — same fonts, sizes, colors throughout
- Fix all spelling errors
- If agenda/index slide is missing or incomplete, complete it with all slide titles
- background_color: hex WITHOUT '#' (e.g. "FFFFFF") or null for transparent
- text_colors: hex WITHOUT '#' (e.g. "2E4057")
- Preserve EVERY piece of data — every table row, every bullet point
- Use "title_content" layout for bullet-point slides
"""


# ── Main Node ─────────────────────────────────────────────────────────────────

async def document_agent_node(state: AgentState) -> dict:
    """
    LangGraph node for document processing.
    Runs extraction + style parsing in parallel, then LLM build, then validation.
    """
    intent = state.get("intent", "doc_task_docx")
    retry_count = state.get("retry_count", 0)
    critic_feedback = state.get("critic_feedback", "")
    source_path = state.get("source_doc_path", "")
    template_path = state.get("template_doc_path", "")
    requirements = state.get("requirements", [])
    session_id = state.get("session_id", "output")

    if not source_path:
        return {
            "draft_response": "No source document found. Please upload a file.",
            "active_agent": "document_agent",
            "requirements_met": False,
        }

    # ── Step 1: Parallel extraction ───────────────────────────────────────────
    logger.info(f"[DocumentAgent] Extracting — intent={intent} retry={retry_count}")

    tasks = [asyncio.to_thread(extract, source_path)]
    if template_path:
        tasks.append(asyncio.to_thread(extract_styles, template_path))

    results = await asyncio.gather(*tasks, return_exceptions=True)

    if isinstance(results[0], Exception):
        logger.error(f"[DocumentAgent] Extraction failed: {results[0]}")
        return {
            "draft_response": f"Failed to extract source document: {results[0]}",
            "active_agent": "document_agent",
            "requirements_met": False,
        }

    extracted_content = results[0]
    extracted_styles = (
        results[1]
        if len(results) > 1 and not isinstance(results[1], Exception)
        else {}
    )

    # ── Step 2: Decide output format ──────────────────────────────────────────
    # Output format is driven by INTENT, NOT source file type.
    # PDF is source-only → output as docx unless explicitly asked for pptx.
    want_pptx = intent == "doc_task_pptx"

    llm = get_llm(task="document")
    output_dir = _get_output_dir(session_id)
    build_mode: str | None = None

    if want_pptx:
        structure = await _llm_build_pptx(
            llm, extracted_content, requirements, critic_feedback, source_path
        )
        output_path = os.path.join(output_dir, "output.pptx")
        if structure:
            # Template priority:
            # 1. Explicit template_doc_path (if provided by user)
            # 2. Source PPTX itself (inherit CGI theme, layouts, design elements)
            # 3. None → blank presentation (fallback)
            pptx_template: str | None = None
            if template_path and template_path.lower().endswith(".pptx"):
                pptx_template = template_path
            elif source_path and source_path.lower().endswith(".pptx"):
                pptx_template = source_path   # use source to inherit CGI theme
            await asyncio.to_thread(build_pptx, structure, output_path, pptx_template)
    elif source_path.lower().endswith(".docx"):
        # DOCX template-mode: modify the SOURCE in place so theme, fonts, colors,
        # cover page and existing tables are all preserved. Apply spelling fixes
        # and (when requested) generate a RACI matrix + process-flow table.
        output_path = os.path.join(output_dir, "output.docx")
        build_mode = "docx_template"
        enhancements = await _llm_build_docx_enhancements(
            llm, extracted_content, requirements, critic_feedback
        )
        structure = enhancements   # non-empty dict satisfies the "built" check
        await asyncio.to_thread(
            build_docx_from_template,
            source_path,
            output_path,
            spelling_fixes=enhancements.get("spelling_fixes") or {},
            raci=enhancements.get("raci"),
            flow=enhancements.get("flow"),
        )
    else:
        # Non-docx source (e.g. PDF) → build a fresh docx from scratch.
        structure = await _llm_build_docx(
            llm, extracted_content, extracted_styles, requirements, critic_feedback
        )
        output_path = os.path.join(output_dir, "output.docx")
        if structure:
            await asyncio.to_thread(build_docx, structure, extracted_styles, output_path)

    if not structure:
        return {
            "draft_response": "Failed to generate document structure from LLM.",
            "active_agent": "document_agent",
            "requirements_met": False,
        }

    # ── Step 3: Validate ──────────────────────────────────────────────────────
    validation = await validate_output(requirements, output_path, build_mode=build_mode)
    requirements_met = validation.get("requirements_met", False)
    feedback = validation.get("feedback", "")

    logger.info(f"[DocumentAgent] requirements_met={requirements_met}")

    if requirements_met:
        draft_response = (
            "Document processed successfully. "
            "All requirements met. Ready to download."
        )
    else:
        failed = validation.get("failed", [])
        draft_response = (
            f"Document built. {len(failed)} requirement(s) flagged: {failed}. "
            f"{feedback}"
        )

    return {
        "extracted_content": extracted_content,
        "extracted_styles": extracted_styles,
        "output_doc_path": output_path,
        "requirements_met": requirements_met,
        "critic_feedback": feedback,
        "retry_count": retry_count + 1,
        "draft_response": draft_response,
        "active_agent": "document_agent",
    }


# ── LLM Build Helpers ─────────────────────────────────────────────────────────

async def _llm_build_docx(
    llm, content: dict, styles: dict, requirements: list, feedback: str
) -> list:
    # Use compact JSON + higher limit (Llama 4 Scout handles 512K tokens)
    content_json = json.dumps(content, separators=(",", ":"))[:20000]
    styles_json  = json.dumps(styles,  separators=(",", ":"))[:4000]

    prompt = f"""Source document content:
{content_json}

Template styles:
{styles_json}

Requirements:
{json.dumps(requirements, indent=2)}"""

    if feedback:
        prompt += f"\n\nPrevious attempt feedback — fix these issues:\n{feedback}"

    messages = [
        SystemMessage(content=DOCX_BUILD_SYSTEM),
        HumanMessage(content=prompt),
    ]
    return await _call_llm_json(llm, messages)


# ── DOCX template-mode enhancements (spelling + RACI + flow) ──────────────────

DOCX_ENHANCE_SYSTEM = """You are a process-analysis specialist for business SOP documents.

From the document provided, extract structured data as JSON. You MUST ground every
value in the document text. Do NOT invent stakeholders, activities, or steps that
are not described in the document. Accuracy and zero hallucination are mandatory.

Return ONLY valid JSON (no preamble, no markdown fences):
{
  "spelling_fixes": { "wrongword": "RightWord", ... },
  "raci": {
    "stakeholders": ["Team A", "Team B", ...],
    "activities": ["1. ...", "2. ...", ...],
    "grid": [ ["R/A","C","-", ...], ... ]
  },
  "flow": {
    "title": "START - ...",
    "steps": [ {"stakeholder": "...", "description": "..."}, ... ]
  }
}

SPELLING_FIXES rules:
- Include ONLY real, unambiguous fixes you can see in the text: typos, inconsistent
  capitalisation of product/brand names (e.g. "Docusign" -> "DocuSign"), obvious
  misspellings. Use exact substrings as they appear. Empty {} if none found.
- Do NOT include proper nouns you are unsure about.

RACI rules (R=Responsible, A=Accountable, C=Consulted, I=Informed):
- "stakeholders": the teams/roles named in the document (e.g. from a "Teams Involved
  & Responsibilities" section). 6-10 columns typical.
- "activities": the major process areas — usually the numbered top-level sections
  (Vendor Empanelment, Governance, Audit, Payment, Termination, etc.).
- "grid": one row per activity, in the same order; one cell per stakeholder, in the
  same order. Each cell is EXACTLY one of: "R", "A", "C", "I", "R/A", "-".
- Assign roles based ONLY on responsibilities described in the document. If a
  stakeholder has no stated role in an activity, use "-".
- Each activity row must have exactly ONE accountable owner — usually expressed as
  "R/A" on the primary owning team.

FLOW rules:
- Reconstruct the end-to-end process as an ordered list of steps.
- Each step: the owning "stakeholder" and a concise "description" (one line).
- Keep the real sequence and ownership from the document. 10-20 steps typical.
- Mark the final step's description with " — END".

If the document clearly has no RACI-relevant content, return raci with empty lists.
If it has no sequential process, return flow with an empty steps list.
"""


def _doc_has_heading(content: dict, keywords: list[str]) -> bool:
    kws = [k.lower() for k in keywords]
    for s in content.get("sections", []):
        h = (s.get("heading") or "").lower()
        if any(k in h for k in kws):
            return True
    return False


def _condense_document(content: dict, per_section_chars: int = 320, total_cap: int = 28000) -> str:
    """
    Build a compact full-document view: every section heading + a snippet of its
    content. This keeps the WHOLE document structure visible to the LLM (so RACI
    activities and flow steps are grounded) while staying within the token budget.
    """
    lines: list[str] = []
    for s in content.get("sections", []):
        heading = (s.get("heading") or "").strip()
        body = (s.get("content") or "").strip().replace("\n", " ")
        if len(body) > per_section_chars:
            body = body[:per_section_chars] + "…"
        if heading or body:
            lines.append(f"## {heading}\n{body}")
    text = "\n".join(lines)
    return text[:total_cap]


async def _llm_build_docx_enhancements(
    llm, content: dict, requirements: list, feedback: str
) -> dict:
    """
    Generate grounded spelling fixes + RACI matrix + process flow for DOCX
    template-mode. Returns {"spelling_fixes": {...}, "raci": {...}, "flow": {...}}.
    Any piece may be empty/None if not applicable or on failure.
    """
    req_text = " ".join(requirements).lower()

    want_raci = (
        any(k in req_text for k in ("raci", "matrix", "responsib", "accountab"))
        or _doc_has_heading(content, ["raci"])
    )
    want_flow = (
        any(k in req_text for k in ("flow", "flowchart", "workflow", "diagram"))
        or _doc_has_heading(content, ["flow", "workflow"])
    )

    doc_view = _condense_document(content)
    all_headings = [s.get("heading", "") for s in content.get("sections", [])]

    prompt = (
        f"Document title: {content.get('filename', 'document')}\n\n"
        f"All section headings (in order):\n{json.dumps(all_headings, ensure_ascii=False)[:6000]}\n\n"
        f"Condensed document content:\n{doc_view}\n\n"
        f"User requirements:\n" + "\n".join(f"- {r}" for r in requirements) + "\n\n"
        f"Generate: spelling_fixes (always), "
        f"{'a RACI matrix' if want_raci else 'raci with empty lists'}, "
        f"{'a process flow' if want_flow else 'flow with empty steps'}."
    )
    if feedback:
        prompt += f"\n\nPrevious feedback to address:\n{feedback}"

    messages = [
        SystemMessage(content=DOCX_ENHANCE_SYSTEM),
        HumanMessage(content=prompt),
    ]

    try:
        response = await llm.ainvoke(messages)
        raw = response.content.strip()
        parsed = _strip_and_parse_json(raw)
        if not isinstance(parsed, dict):
            logger.warning("[DocumentAgent] DOCX enhancements: non-dict result")
            return {"spelling_fixes": {}, "raci": None, "flow": None}

        # Normalise / guard
        result = {
            "spelling_fixes": parsed.get("spelling_fixes") or {},
            "raci": None,
            "flow": None,
        }
        raci = parsed.get("raci") or {}
        if want_raci and raci.get("activities") and raci.get("stakeholders"):
            result["raci"] = raci
        flow = parsed.get("flow") or {}
        if want_flow and flow.get("steps"):
            result["flow"] = flow

        logger.info(
            f"[DocumentAgent] DOCX enhancements: fixes={len(result['spelling_fixes'])} "
            f"raci={'yes' if result['raci'] else 'no'} flow={'yes' if result['flow'] else 'no'}"
        )
        return result

    except Exception as e:
        logger.error(f"[DocumentAgent] DOCX enhancements error: {e}")
        return {"spelling_fixes": {}, "raci": None, "flow": None}


async def _llm_build_pptx(
    llm,
    content: dict,
    requirements: list,
    feedback: str,
    source_path: str,
) -> list:
    """
    Build PPTX structure using PER-SLIDE LLM processing.

    WHY PER-SLIDE:
    Dumping the whole deck as one JSON prompt causes silent truncation.
    A 7-slide deck with tables is ~30,000 chars; the old [:6000] limit
    meant slides 4-7 were never seen by the LLM → hallucinated empty slides.

    Per-slide processing: each slide gets its own LLM call with its full
    content (no truncation). 7 slides = 7 calls, each small and focused.
    """
    req_text = " ".join(requirements).lower()

    # ── Extract theme ─────────────────────────────────────────────────────────
    theme: dict = {}
    if source_path.lower().endswith(".pptx"):
        slide_match = re.search(r'slide\s+(\d+)', req_text)
        try:
            slide_index = (int(slide_match.group(1)) - 1) if slide_match else 3
            theme = extract_pptx_theme(source_path, slide_index=slide_index)
            logger.info(f"[DocumentAgent] Theme from slide {slide_index + 1}: "
                        f"fonts={theme.get('fonts')} master={theme.get('master_fonts')}")
        except Exception as e:
            logger.warning(f"[DocumentAgent] Theme extraction failed: {e}")

    # Build a compact, LLM-ready theme string
    theme_str = _build_theme_instructions(theme)

    # ── Get all source slides ─────────────────────────────────────────────────
    sections = content.get("sections", [])
    if not sections:
        logger.error("[DocumentAgent] No sections in extracted content")
        return []

    # Collect all slide titles for the agenda slide
    all_titles = [s.get("heading", f"Slide {s.get('slide_number', i+1)}")
                  for i, s in enumerate(sections)]

    # ── Process each slide individually ───────────────────────────────────────
    all_slides: list[dict] = []
    requirements_text = "\n".join(f"- {r}" for r in requirements)

    # Semaphore limits concurrent LLM calls to avoid 429 rate-limit errors.
    # Groq free tier: 30K TPM with ~2-4K tokens/slide → max 3 concurrent safe.
    _sem = asyncio.Semaphore(3)

    tasks = [
        _llm_process_single_slide(
            llm, section, all_titles, theme_str, requirements_text, feedback, _sem
        )
        for section in sections
    ]

    # Run slides concurrently (semaphore-gated to respect TPM limit)
    results = await asyncio.gather(*tasks, return_exceptions=True)

    canon_theme = _default_theme(theme)   # always-correct theme for enforcement

    for i, result in enumerate(results):
        if isinstance(result, Exception):
            logger.error(f"[DocumentAgent] Slide {i+1} LLM error: {result}")
            all_slides.append({
                "layout": "title_content",
                "title": sections[i].get("heading", f"Slide {i+1}"),
                "content": "(Content generation failed)",
                "theme": canon_theme,
            })
        elif result:
            # Enforce extracted theme values — LLM sometimes drifts to black text
            result["theme"] = _enforce_theme(result.get("theme") or {}, canon_theme)
            all_slides.append(result)
        else:
            logger.warning(f"[DocumentAgent] Slide {i+1} returned empty — using fallback")
            all_slides.append({
                "layout": "title_content",
                "title": sections[i].get("heading", f"Slide {i+1}"),
                "content": "",
                "theme": canon_theme,
            })

    logger.info(f"[DocumentAgent] Built {len(all_slides)} slides via per-slide processing")
    return all_slides


PPTX_SINGLE_SLIDE_SYSTEM = """You are a presentation design specialist.

You will receive ONE slide's content from a source presentation, plus the design theme to apply.

Your job: return a SINGLE JSON OBJECT (not a list) for this one slide.

CRITICAL RULES:
1. Preserve ALL content — every table row, every bullet point. Do NOT summarise or drop data.
2. If the slide has a "tables" key, reproduce it EXACTLY in your output "table" field. Do NOT convert tables to bullets.
3. Fix any spelling errors in the content.
4. Apply the provided theme (fonts, colors, sizes) in the theme field of your output.
5. Return ONLY the JSON object. No preamble, no markdown fences.

Output format:
{
  "layout": "title" | "title_content" | "blank",
  "title": "...",
  "content": "string or list of bullet strings",   // use for text-only slides
  "table": {                                         // use INSTEAD OF content for table slides
    "headers": ["Col1", "Col2", ...],
    "rows": [["R1C1", "R1C2"], ...]
  },
  "notes": "",
  "theme": {
    "fonts": ["Title Font", "Body Font"],
    "font_sizes": [title_pt, body_pt],
    "text_colors": ["RRGGBB_title", "RRGGBB_body"],
    "background_color": "RRGGBB or null"
  }
}

LAYOUT RULES:
- Cover / title slides (slide 1) → "title"
- Slides with ONLY bullets/text (no table) → "title_content"
- Slides with a TABLE → "title_content"  ← IMPORTANT: use title_content, NOT blank
- "blank" is for rare cases with no title at all

FONT SIZE RULES:
- Cover slide title (layout=title): use 36pt for title
- Section/content slide title (layout=title_content): use 28pt for title
- Body text: use 18pt
- The table font size is handled automatically by the builder; just set body_pt=18

AGENDA SLIDE RULES (when slide title contains "Agenda"):
- Do NOT list every slide title literally. Instead identify 2-4 meaningful SECTIONS.
- Group slides into logical sections, e.g. "Active Recruitment Vendors", "Active Non-Recruitment Vendors", "Vendor Pipeline – RFP"
- Use the all_titles list to understand the deck structure, then create clean section names.
- Format as a list of CLEAN section names WITHOUT numbers: ["Active Recruitment Vendors", "Active Non-Recruitment Vendors", "Vendor Pipeline – RFP"]
- Do NOT prefix items with "01." or "1." — the numbers are handled automatically by the slide layout

SPELLING: Fix obvious errors (wrong capitalisation, typos). Preserve proper nouns.
"""


async def _llm_process_single_slide(
    llm,
    section: dict,
    all_titles: list[str],
    theme_str: str,
    requirements_text: str,
    feedback: str,
    semaphore: asyncio.Semaphore | None = None,
) -> dict | None:
    """
    Process a single slide section through the LLM.
    Retries up to 3 times on 429 rate-limit errors with exponential backoff.
    Returns a single slide dict, or None on failure.
    """
    slide_num = section.get("slide_number", "?")
    heading   = section.get("heading", "")
    content   = section.get("content", "")
    tables    = section.get("tables", [])

    # Compact JSON for this single slide — no truncation needed, each slide is small
    slide_data = {
        "slide_number": slide_num,
        "title": heading,
        "content": content,
    }
    if tables:
        slide_data["tables"] = tables

    slide_json = json.dumps(slide_data, separators=(",", ":"))

    is_cover  = (slide_num == 1)
    is_agenda = "agenda" in heading.lower()

    prompt = (
        f"Slide {slide_num} of {len(all_titles)} — "
        f"{'COVER SLIDE (use layout=title, title font 36pt)' if is_cover else 'AGENDA SLIDE (apply agenda rules above)' if is_agenda else 'CONTENT SLIDE'}"
        f" source content:\n{slide_json}\n\n"
        f"All slide titles in deck (for agenda context + section grouping):\n"
        f"{json.dumps(all_titles, separators=(',', ':'))}\n\n"
        f"Design theme to apply:\n{theme_str}\n\n"
        f"Requirements:\n{requirements_text}"
    )
    if feedback:
        prompt += f"\n\nPrevious feedback to fix:\n{feedback}"

    messages = [
        SystemMessage(content=PPTX_SINGLE_SLIDE_SYSTEM),
        HumanMessage(content=prompt),
    ]

    max_retries = 4
    for attempt in range(max_retries):
        try:
            # Acquire semaphore to cap concurrent calls and stay within TPM limit
            if semaphore:
                async with semaphore:
                    response = await llm.ainvoke(messages)
            else:
                response = await llm.ainvoke(messages)

            raw    = response.content.strip()
            parsed = _strip_and_parse_json(raw)

            if isinstance(parsed, dict):
                logger.info(
                    f"[DocumentAgent] Slide {slide_num} OK: "
                    f"layout={parsed.get('layout')} "
                    f"table={'yes' if parsed.get('table') else 'no'}"
                )
                return parsed
            elif isinstance(parsed, list) and parsed:
                logger.warning(f"[DocumentAgent] Slide {slide_num} returned list, taking [0]")
                return parsed[0]
            else:
                logger.error(f"[DocumentAgent] Slide {slide_num}: unexpected JSON shape {type(parsed)}")
                return None

        except Exception as e:
            err_str = str(e)
            is_rate_limit = "429" in err_str or "rate_limit" in err_str.lower()

            if is_rate_limit and attempt < max_retries - 1:
                # Parse retry-after from error if available, else exponential backoff
                wait = _parse_retry_after(err_str) or (2 ** attempt * 3)  # 3s, 6s, 12s
                logger.warning(
                    f"[DocumentAgent] Slide {slide_num} rate-limited "
                    f"(attempt {attempt + 1}/{max_retries}), retrying in {wait:.1f}s"
                )
                await asyncio.sleep(wait)
                continue

            logger.error(f"[DocumentAgent] Slide {slide_num} error: {e}")
            return None

    logger.error(f"[DocumentAgent] Slide {slide_num}: all {max_retries} retries exhausted")
    return None


def _parse_retry_after(err_str: str) -> float | None:
    """
    Try to extract a wait time from a Groq 429 error message.
    Examples: 'Please try again in 457.999999ms' → 0.458
              'Please try again in 3.768s' → 3.768
    """
    import re
    m = re.search(r"try again in ([\d.]+)(ms|s)", err_str, re.IGNORECASE)
    if not m:
        return None
    val  = float(m.group(1))
    unit = m.group(2).lower()
    wait = val / 1000 if unit == "ms" else val
    return max(wait + 0.5, 1.0)   # add 0.5s buffer; minimum 1s


def _build_theme_instructions(theme: dict) -> str:
    """Convert extracted theme dict into a compact instruction string for the LLM."""
    if not theme:
        return (
            "No theme data — use: fonts=[Calibri, Calibri], font_sizes=[28, 14], "
            "text_colors=[1F3864, 333333], background_color=null"
        )

    master = theme.get("master_fonts", {})
    accent = theme.get("theme_accent_colors", {})
    fonts  = list(theme.get("fonts", []))
    sizes  = list(theme.get("font_sizes", []))
    colors = list(theme.get("text_colors", []))
    bg     = theme.get("background_color")

    # Resolve fonts from master if run-level fonts are empty
    if not fonts:
        title_font = master.get("title") or "Calibri"
        body_font  = master.get("body")  or "Calibri"
        fonts = [title_font, body_font]

    # Resolve colors: prefer run-level, then accent colors from theme
    # Convention: colors[0]=title/primary, colors[1]=body
    # dk2 is the brand dark tone (navy/title), dk1 is black (body text)
    if not colors:
        dk2 = accent.get("dk2") or "1F3864"
        dk1 = accent.get("dk1") or "333333"
        colors = [dk2, dk1]

    if not sizes:
        sizes = [28, 18]
    elif len(sizes) == 1:
        sizes.append(18)   # body size when only title size extracted

    lines = [
        f"fonts: {fonts}",
        f"font_sizes: {sizes} (pt)",
        f"text_colors: {colors} (hex, no #)",
        f"background_color: {bg or 'null'} (null=transparent/inherit)",
        f"master_title_font: {master.get('title', 'N/A')}",
        f"master_body_font: {master.get('body', 'N/A')}",
        f"theme_accent_colors: {accent}",
        "IMPORTANT: Use these exact values in every slide's theme object.",
        "If text_colors is empty, use dk1 and dk2 from theme_accent_colors.",
        "If fonts is empty, use master_title_font and master_body_font.",
    ]
    return "\n".join(lines)


def _default_theme(theme: dict) -> dict:
    """Return a safe default theme dict from extracted theme or hardcoded fallback."""
    master = theme.get("master_fonts", {})
    accent = theme.get("theme_accent_colors", {})

    fonts  = list(theme.get("fonts") or [])
    sizes  = list(theme.get("font_sizes") or [28, 18])
    colors = list(theme.get("text_colors") or [])
    bg     = theme.get("background_color")

    # Resolve fonts: run-level → master → hardcoded
    if not fonts or not fonts[0] or fonts[0] == "inherit":
        fonts = [master.get("title") or "Calibri"]
    if len(fonts) < 2 or not fonts[1] or fonts[1] == "inherit":
        fonts.append(master.get("body") or fonts[0])

    # Resolve colors: run-level → accent scheme → hardcoded
    if not colors:
        dk1 = accent.get("dk1") or "000000"
        dk2 = accent.get("dk2") or "1F3864"
        colors = [dk2, dk1]   # dk2 is the brand dark color (title/accent), dk1 is black (body)

    # Ensure we always have two sizes: [title_pt, body_pt]
    if len(sizes) == 1:
        sizes.append(18)   # standard body text size when only title size is known

    return {
        "fonts": fonts[:2],
        "font_sizes": sizes[:2],
        "text_colors": colors[:2],
        "background_color": bg,
    }


def _enforce_theme(llm_theme: dict, canon: dict) -> dict:
    """
    Merge the LLM's returned theme with the canonical extracted theme.
    The LLM sometimes drifts to black text or forgets fonts — this corrects it.

    Rules:
    - fonts: use LLM value only if non-empty and not 'inherit'; else use canon
    - font_sizes: use LLM value only if both sizes are present and reasonable (>6pt)
    - text_colors: use LLM value only if title color is NOT black ('000000')
    - background_color: keep LLM value (it may intentionally set a bg)
    """
    # --- fonts ---
    llm_fonts = [f for f in (llm_theme.get("fonts") or []) if f and f != "inherit"]
    fonts = llm_fonts[:2] if len(llm_fonts) >= 2 else canon["fonts"]

    # --- font_sizes ---
    llm_sizes = llm_theme.get("font_sizes") or []
    if len(llm_sizes) >= 2 and all(isinstance(s, (int, float)) and s > 6 for s in llm_sizes[:2]):
        sizes = llm_sizes[:2]
    else:
        sizes = canon["font_sizes"]

    # --- text_colors ---
    llm_colors = llm_theme.get("text_colors") or []
    canon_colors = canon["text_colors"]
    BLACK = "000000"
    # Only accept LLM title color if it's not pure black or empty
    if llm_colors and llm_colors[0] and llm_colors[0].upper().lstrip("#") != BLACK:
        title_color = llm_colors[0].lstrip("#")
    else:
        title_color = canon_colors[0]
    # Body color: accept LLM value (black body text is fine)
    if len(llm_colors) > 1 and llm_colors[1]:
        body_color = llm_colors[1].lstrip("#")
    else:
        body_color = canon_colors[1] if len(canon_colors) > 1 else BLACK

    return {
        "fonts": fonts,
        "font_sizes": sizes,
        "text_colors": [title_color, body_color],
        "background_color": llm_theme.get("background_color") or canon.get("background_color"),
    }


# ── JSON parsing utility ──────────────────────────────────────────────────────

async def _call_llm_json(llm, messages) -> list:
    """Call LLM and parse JSON response. Returns empty list on failure."""
    try:
        response = await llm.ainvoke(messages)
        raw = response.content.strip()
        parsed = _strip_and_parse_json(raw)

        # Accept both list and {"slides": [...]} / {"blocks": [...]}
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            for key in ("slides", "blocks", "content", "structure"):
                if key in parsed and isinstance(parsed[key], list):
                    return parsed[key]

        logger.error(f"[DocumentAgent] Unexpected JSON shape: {type(parsed)}")
        return []

    except Exception as e:
        logger.error(f"[DocumentAgent] LLM parse error: {e}")
        return []


def _strip_and_parse_json(raw: str):
    """
    Strip markdown code fences and DeepSeek/Qwen thinking tags, then parse JSON.
    Falls back to bracket-matching extraction when direct parse fails.

    Handles:
    - <think>...</think> tags (DeepSeek R1, Qwen QwQ)
    - ```json ... ``` fences
    - Stray text before or after the JSON object/array
    - Truncated JSON (raises JSONDecodeError — caller falls back)
    """
    import json
    import re

    # ── 1. Strip <think>...</think> tags ──────────────────────────────────────
    if "<think>" in raw:
        raw = raw.split("</think>")[-1].strip()

    # ── 2. Strip ``` fences ───────────────────────────────────────────────────
    if "```" in raw:
        m = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", raw)
        if m:
            raw = m.group(1).strip()
        else:
            # Malformed fence: strip leading ``` and try
            raw = raw.replace("```json", "").replace("```", "").strip()

    # ── 3. Direct parse (fastest path) ───────────────────────────────────────
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # ── 4. Bracket-matching extraction ────────────────────────────────────────
    # LLM sometimes emits text before/after the JSON. Find the outermost
    # { ... } or [ ... ] and parse that.
    for open_ch, close_ch in (('{', '}'), ('[', ']')):
        start = raw.find(open_ch)
        if start == -1:
            continue
        depth       = 0
        in_string   = False
        escape_next = False
        for i, ch in enumerate(raw[start:], start):
            if escape_next:
                escape_next = False
                continue
            if ch == '\\' and in_string:
                escape_next = True
                continue
            if ch == '"':
                in_string = not in_string
            elif not in_string:
                if ch == open_ch:
                    depth += 1
                elif ch == close_ch:
                    depth -= 1
                    if depth == 0:
                        candidate = raw[start : i + 1]
                        try:
                            return json.loads(candidate)
                        except json.JSONDecodeError:
                            break   # found brackets but still invalid; try next

    # ── 5. Last resort: re-raise with original error ──────────────────────────
    return json.loads(raw)
