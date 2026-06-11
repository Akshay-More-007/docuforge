"""
engine.py — Learn-mode content engine.

Turns a document (or pasted notes) into an interactive study session:
  - quiz        → multiple-choice questions with explanations
  - flashcards  → front/back cards with hints
  - interview   → open questions with key points + model answers, LLM-graded

Design: ONE grounded LLM call generates the whole item set up front, so the
game loop itself is instant (no per-click LLM latency). Only interview-answer
grading calls the LLM during play (small, fast call).

Every item carries a `source_ref` naming the section/slide it came from, so
wrong answers can be aggregated into "weak areas" for review.
"""

import asyncio
import json
import logging

from src.document.extractor import extract
from src.llm.router import get_llm
from src.agents.document_agent import _strip_and_parse_json, _response_text

logger = logging.getLogger(__name__)

MODES = ("quiz", "flashcards", "interview")

# ── Generation prompts ────────────────────────────────────────────────────────

GEN_SYSTEM = """You are an expert learning-content designer and exam writer.

You receive study material (sections of a document). Create study items that are
GROUNDED ONLY in that material — never invent facts that are not in the text.
Mix recall questions with applied/scenario questions so the learner truly has to
understand the material, not just memorise words.

Return ONLY strict valid JSON (double quotes, no trailing commas, no markdown
fences, no preamble): {"items": [ ... ]}

Item schema depends on the requested mode:

MODE quiz — each item:
{
  "question": "clear, self-contained question",
  "options": ["A", "B", "C", "D"],          // exactly 4, one correct
  "answer_index": 0,                         // index of the correct option
  "explanation": "why the answer is right, citing the document fact",
  "source_ref": "section/slide the item came from",
  "difficulty": "easy" | "medium" | "hard"
}
Distractors must be PLAUSIBLE (drawn from related document content), never
joke options. Vary the position of the correct answer across items.

MODE flashcards — each item:
{
  "front": "term, concept or question",
  "back": "concise answer/definition from the document",
  "hint": "one short nudge that doesn't give the answer away",
  "source_ref": "section/slide"
}

MODE interview — each item:
{
  "question": "open interview-style question about the material",
  "key_points": ["3-5 facts a strong answer must mention"],
  "model_answer": "a strong 3-6 sentence answer",
  "source_ref": "section/slide",
  "difficulty": "easy" | "medium" | "hard"
}
Make interview questions the kind a real interviewer would ask about this
material: process questions, "what would you do if", "why is X done before Y".

GENERAL RULES:
- Cover DIFFERENT sections of the material — don't cluster on one section.
- Respect the requested difficulty; "mixed" = a spread of easy→hard.
- If a focus topic is given, weight items toward it (but stay grounded).
- Write in clear professional English regardless of source-document typos.
"""

EVAL_SYSTEM = """You are a fair, encouraging interview coach grading one answer.

You receive: the question, the key points a strong answer should mention, a
model answer, and the candidate's answer.

Return ONLY strict JSON:
{
  "score": 0-10,                 // 10 = covers all key points accurately
  "feedback": "2-3 sentences: what was good + what to improve, friendly tone",
  "missed_points": ["key points the answer did not cover"],   // [] if none
  "covered_points": ["key points the answer did cover"]
}

Scoring guide: every key point clearly covered ≈ proportional share of 10;
factually WRONG statements cost extra. An empty or off-topic answer scores 0-1.
Judge meaning, not wording — synonyms and paraphrases count as covered.
"""


# ── Material preparation ──────────────────────────────────────────────────────

def _condense_for_learning(content: dict, per_section_chars: int = 900,
                           total_cap: int = 24000) -> str:
    """Heading + generous snippet per section — enough substance to write
    grounded questions while staying inside the token budget."""
    lines: list[str] = []
    for s in content.get("sections", []):
        heading = (s.get("heading") or "").strip()
        body = (s.get("content") or "").strip().replace("\n", " ")
        # include table rows when present (slides/tables hold the real facts)
        for tbl in s.get("tables", []) or []:
            rows = tbl.get("all_rows") or []
            body += " " + " ; ".join(" | ".join(r) for r in rows[:12])
        if len(body) > per_section_chars:
            body = body[:per_section_chars] + "…"
        if heading or body:
            lines.append(f"## {heading}\n{body}")
    return "\n".join(lines)[:total_cap]


def _material_from_text(raw_text: str, title: str = "Pasted notes") -> dict:
    """Wrap pasted notes in the extractor's content shape."""
    return {
        "filename": title,
        "sections": [{"heading": title, "content": raw_text.strip()}],
    }


# ── Generation ────────────────────────────────────────────────────────────────

async def generate_session(
    *,
    source_path: str | None = None,
    raw_text: str | None = None,
    mode: str = "quiz",
    n_items: int = 8,
    difficulty: str = "mixed",
    focus: str = "",
) -> dict:
    """
    Generate a study session from a document path OR pasted text.

    Returns {"mode", "doc_title", "items": [...]} — items validated per mode.
    Raises ValueError when generation fails or yields too few usable items.
    """
    if mode not in MODES:
        raise ValueError(f"Unknown mode: {mode}")

    if source_path:
        content = await asyncio.to_thread(extract, source_path)
    elif raw_text and raw_text.strip():
        content = _material_from_text(raw_text)
    else:
        raise ValueError("Provide a document or pasted notes to learn from.")

    material = _condense_for_learning(content)
    if len(material) < 200:
        raise ValueError("Not enough material to build a session from.")

    doc_title = content.get("filename", "your material")
    prompt = (
        f"MODE: {mode}\n"
        f"Number of items: {n_items}\n"
        f"Difficulty: {difficulty}\n"
        f"Focus topic: {focus or '(none — cover the whole material)'}\n\n"
        f"Study material from \"{doc_title}\":\n{material}"
    )

    llm = get_llm(task="learn")   # Gemini chain (lite-first) → Groq fallback
    from langchain_core.messages import HumanMessage, SystemMessage
    messages = [SystemMessage(content=GEN_SYSTEM), HumanMessage(content=prompt)]

    last_err: Exception | None = None
    for attempt in range(3):
        try:
            response = await llm.ainvoke(messages)
            parsed = _strip_and_parse_json(_response_text(response))
            items = _validate_items(parsed, mode)
            if len(items) >= min(3, n_items):
                logger.info(f"[Learn] Generated {len(items)} {mode} item(s) "
                            f"from {doc_title!r}")
                return {"mode": mode, "doc_title": doc_title,
                        "items": items[:n_items]}
            last_err = ValueError(f"only {len(items)} usable items")
        except Exception as e:                          # JSON/transport slip
            last_err = e
            logger.warning(f"[Learn] generation attempt {attempt + 1} failed: {e}")
            await asyncio.sleep(1.5 * attempt)
    raise ValueError(f"Could not generate a {mode} session: {last_err}")


def _validate_items(parsed, mode: str) -> list[dict]:
    """Keep only structurally sound items for the given mode."""
    if isinstance(parsed, list):
        raw_items = parsed
    elif isinstance(parsed, dict):
        raw_items = parsed.get("items") or []
    else:
        return []

    items: list[dict] = []
    for it in raw_items:
        if not isinstance(it, dict):
            continue
        try:
            if mode == "quiz":
                opts = it.get("options")
                idx = it.get("answer_index")
                if (it.get("question") and isinstance(opts, list)
                        and len(opts) == 4 and isinstance(idx, int)
                        and 0 <= idx <= 3):
                    items.append({
                        "question": str(it["question"]),
                        "options": [str(o) for o in opts],
                        "answer_index": idx,
                        "explanation": str(it.get("explanation", "")),
                        "source_ref": str(it.get("source_ref", "")),
                        "difficulty": str(it.get("difficulty", "medium")),
                    })
            elif mode == "flashcards":
                if it.get("front") and it.get("back"):
                    items.append({
                        "front": str(it["front"]),
                        "back": str(it["back"]),
                        "hint": str(it.get("hint", "")),
                        "source_ref": str(it.get("source_ref", "")),
                    })
            elif mode == "interview":
                kps = it.get("key_points")
                if it.get("question") and isinstance(kps, list) and kps:
                    items.append({
                        "question": str(it["question"]),
                        "key_points": [str(k) for k in kps],
                        "model_answer": str(it.get("model_answer", "")),
                        "source_ref": str(it.get("source_ref", "")),
                        "difficulty": str(it.get("difficulty", "medium")),
                    })
        except Exception:
            continue
    return items


# ── Interview answer grading ──────────────────────────────────────────────────

async def evaluate_interview_answer(item: dict, user_answer: str) -> dict:
    """Grade a free-text answer. Always returns a usable dict (never raises)."""
    user_answer = (user_answer or "").strip()
    if not user_answer:
        return {"score": 0, "feedback": "No answer given — try putting the key "
                "ideas in your own words, even briefly.",
                "missed_points": list(item.get("key_points", [])),
                "covered_points": []}

    prompt = (
        f"QUESTION:\n{item['question']}\n\n"
        f"KEY POINTS a strong answer must mention:\n"
        + "\n".join(f"- {k}" for k in item.get("key_points", []))
        + f"\n\nMODEL ANSWER:\n{item.get('model_answer', '(none)')}\n\n"
        f"CANDIDATE'S ANSWER:\n{user_answer}"
    )

    llm = get_llm(task="chat")   # small fast call (Groq)
    from langchain_core.messages import HumanMessage, SystemMessage
    try:
        response = await llm.ainvoke(
            [SystemMessage(content=EVAL_SYSTEM), HumanMessage(content=prompt)]
        )
        result = _strip_and_parse_json(_response_text(response))
        score = result.get("score", 0)
        score = max(0, min(10, int(score)))
        return {
            "score": score,
            "feedback": str(result.get("feedback", "")),
            "missed_points": [str(m) for m in result.get("missed_points") or []],
            "covered_points": [str(c) for c in result.get("covered_points") or []],
        }
    except Exception as e:
        logger.error(f"[Learn] evaluation failed: {e}")
        return {"score": 0,
                "feedback": "Grading hit a temporary error — compare your answer "
                            "with the model answer below.",
                "missed_points": [], "covered_points": []}


# ── Voice: speech-to-text for spoken interview answers ───────────────────────

def transcribe_audio(audio_bytes: bytes, filename: str = "answer.wav") -> str:
    """
    Transcribe a recorded answer with Groq Whisper (whisper-large-v3-turbo —
    fast and free-tier friendly). Returns "" on failure rather than raising,
    so the UI can fall back to typed input.
    """
    import os
    try:
        from groq import Groq
        client = Groq(api_key=os.getenv("GROQ_API_KEY"))
        result = client.audio.transcriptions.create(
            file=(filename, audio_bytes),
            model=os.getenv("GROQ_MODEL_STT", "whisper-large-v3-turbo"),
            language="en",
            temperature=0.0,
        )
        text = (result.text or "").strip()
        logger.info(f"[Learn] Transcribed {len(audio_bytes)} bytes → "
                    f"{len(text)} chars")
        return text
    except Exception as e:
        logger.error(f"[Learn] transcription failed: {e}")
        return ""


# ── Result helpers (pure, unit-testable) ──────────────────────────────────────

def grade_label(pct: float) -> tuple[str, str]:
    """Return (label, badge_class) for a final quiz score percentage."""
    if pct >= 90:
        return "Outstanding — interview ready", "badge-success"
    if pct >= 75:
        return "Strong — minor gaps", "badge-success"
    if pct >= 50:
        return "Getting there — review the weak areas", "badge-info"
    return "Needs review — go through the material again", "badge-error"


def weak_areas(answers: list[dict], top_n: int = 3) -> list[tuple[str, int]]:
    """Aggregate wrong answers by source_ref → [(section, misses)] descending."""
    counts: dict[str, int] = {}
    for a in answers:
        if not a.get("correct"):
            ref = (a.get("source_ref") or "General").strip() or "General"
            counts[ref] = counts.get(ref, 0) + 1
    return sorted(counts.items(), key=lambda kv: -kv[1])[:top_n]
