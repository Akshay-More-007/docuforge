"""
learn.py — Learn mode: turn a document into an interactive study game.

Modes: Quiz (MCQ, instant feedback) · Flashcards (flip + self-grade, missed
cards loop back) · Interview Prep (free-text answers graded by the LLM).

Phases (st.session_state.learn_phase):
  setup → playing → done   (+ "review" for going through quiz mistakes)

All gameplay is local session state — only generation (once, up front) and
interview grading call the LLM, via the same isolated-thread pattern chat uses.
"""

import asyncio
import concurrent.futures

import streamlit as st

from src.auth.guards import require_auth
from src.app.components.sidebar import render_sidebar
from src.app.compat import rerun
from src.learn.engine import (
    generate_session, evaluate_interview_answer, grade_label, weak_areas,
    transcribe_audio,
)

MODE_CARDS = [
    ("quiz", "🎯", "Quiz", "Multiple choice with instant feedback and explanations."),
    ("flashcards", "🃏", "Flashcards", "Flip cards, self-grade — missed cards come back."),
    ("interview", "🎤", "Interview Prep", "Answer in your own words, get coached and scored."),
]


def _run_async(coro):
    """Run a coroutine on a dedicated thread/event loop (Streamlit-safe)."""
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
        return pool.submit(asyncio.run, coro).result()


def _init_state():
    defaults = {
        "learn_phase": "setup",
        "learn_mode": "quiz",
        "learn_items": [],
        "learn_idx": 0,
        "learn_answers": [],       # per-item result dicts
        "learn_answered": None,    # quiz: chosen option index for current item
        "learn_flipped": False,    # flashcards
        "learn_again": [],         # flashcards: missed items to repeat
        "learn_eval": None,        # interview: evaluation of current answer
        "learn_doc_title": "",
        "learn_settings": {"n_items": 8, "difficulty": "mixed", "focus": ""},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_learn():
    require_auth()
    _init_state()
    render_sidebar()

    st.markdown(
        "<div style='padding: 24px 0 8px 0;'>"
        "<span class='brand' style='font-size:16px;'>🎓 Learn</span>"
        "<span style='font-size:11px;color:#666;margin-left:10px;'>"
        "study · polish skills · interview prep</span></div>",
        unsafe_allow_html=True,
    )

    phase = st.session_state.learn_phase
    if phase == "setup":
        _render_setup()
    elif phase == "playing":
        mode = st.session_state.learn_mode
        if mode == "quiz":
            _render_quiz()
        elif mode == "flashcards":
            _render_flashcards()
        else:
            _render_interview()
    elif phase == "review":
        _render_review()
    else:
        _render_done()


# ── Setup ─────────────────────────────────────────────────────────────────────

def _render_setup():
    source_path = st.session_state.get("source_doc_path")
    source_name = st.session_state.get("source_doc_name", "your document")

    if source_path:
        st.markdown(
            f"<div class='doc-card'><div class='doc-name'>📄 {source_name}</div>"
            f"<div style='font-size:11px;color:#666;'>Learning from this document</div></div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<p style='font-size:13px;color:#888;'>Upload a document in the "
            "sidebar — or paste your notes below — and turn it into a game.</p>",
            unsafe_allow_html=True,
        )

    pasted = ""
    if not source_path:
        with st.expander("📝 Paste notes instead", expanded=False):
            pasted = st.text_area(
                "Paste any study material",
                key="learn_paste",
                height=160,
                placeholder="Paste lecture notes, a job description, an article…",
            )

    # Mode cards
    st.markdown("<p class='learn-section-label'>CHOOSE A MODE</p>",
                unsafe_allow_html=True)
    cols = st.columns(len(MODE_CARDS))
    selected = st.session_state.learn_mode
    for col, (mode, icon, title, desc) in zip(cols, MODE_CARDS):
        with col:
            active = "learn-mode-active" if mode == selected else ""
            st.markdown(
                f"<div class='learn-mode-card {active}'>"
                f"<div style='font-size:22px;'>{icon}</div>"
                f"<div class='doc-name'>{title}</div>"
                f"<div style='font-size:11px;color:#888;line-height:1.5;'>{desc}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
            if st.button(("✓ Selected" if mode == selected else "Select"),
                         key=f"mode_{mode}", use_container_width=True,
                         type="primary" if mode == selected else "secondary"):
                st.session_state.learn_mode = mode
                rerun()

    # Settings
    st.markdown("<p class='learn-section-label'>SESSION SETTINGS</p>",
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        n_items = st.select_slider("Questions", options=[5, 8, 10, 15], value=8)
    with c2:
        difficulty = st.selectbox("Difficulty",
                                  ["mixed", "easy", "medium", "hard"], index=0)
    with c3:
        focus = st.text_input("Focus on (optional)",
                              placeholder="e.g. payment process, KPIs, security…")

    st.session_state.learn_settings = {
        "n_items": n_items, "difficulty": difficulty, "focus": focus,
    }

    if st.button("🚀  Start session", type="primary", use_container_width=True):
        if not source_path and not (pasted or "").strip():
            st.warning("Upload a document in the sidebar or paste some notes first.")
            return
        _start_session(source_path, pasted)


def _start_session(source_path: str | None, pasted: str):
    s = st.session_state.learn_settings
    try:
        with st.spinner("Building your session — reading the material and "
                        "writing questions…"):
            session = _run_async(generate_session(
                source_path=source_path,
                raw_text=None if source_path else pasted,
                mode=st.session_state.learn_mode,
                n_items=s["n_items"],
                difficulty=s["difficulty"],
                focus=s["focus"],
            ))
    except Exception as e:
        st.error(f"Couldn't build the session: {e}")
        return

    st.session_state.learn_items = session["items"]
    st.session_state.learn_doc_title = session["doc_title"]
    st.session_state.learn_idx = 0
    st.session_state.learn_answers = []
    st.session_state.learn_answered = None
    st.session_state.learn_flipped = False
    st.session_state.learn_again = []
    st.session_state.learn_eval = None
    st.session_state.learn_phase = "playing"
    rerun()


# ── Shared bits ───────────────────────────────────────────────────────────────

def _progress_header(label: str):
    items = st.session_state.learn_items
    idx = st.session_state.learn_idx
    total = len(items) + len(st.session_state.learn_again)
    st.progress(min(idx / max(total, 1), 1.0))
    st.markdown(
        f"<p style='font-size:12px;color:#888;'>{label} "
        f"<span style='color:#f5a623;'>{min(idx + 1, total)}</span> of {total}"
        f" &nbsp;·&nbsp; {st.session_state.learn_doc_title}</p>",
        unsafe_allow_html=True,
    )


def _question_card(text: str, badge: str = ""):
    badge_html = (f"<span class='badge badge-info' style='float:right;'>{badge}</span>"
                  if badge else "")
    st.markdown(
        f"<div class='learn-question'>{badge_html}{text}</div>",
        unsafe_allow_html=True,
    )


def _end_session_button():
    if st.button("✕ End session", key="end_session"):
        st.session_state.learn_phase = "done"
        rerun()


# ── Quiz ──────────────────────────────────────────────────────────────────────

def _render_quiz():
    items = st.session_state.learn_items
    idx = st.session_state.learn_idx
    if idx >= len(items):
        st.session_state.learn_phase = "done"
        rerun()
        return

    q = items[idx]
    score = sum(1 for a in st.session_state.learn_answers if a.get("correct"))
    _progress_header("Question")
    st.markdown(
        f"<p style='font-size:12px;color:#888;'>Score: "
        f"<span style='color:#2ecc71;'>{score}</span></p>",
        unsafe_allow_html=True,
    )
    _question_card(q["question"], q.get("difficulty", ""))

    answered = st.session_state.learn_answered

    if answered is None:
        for i, opt in enumerate(q["options"]):
            if st.button(f"{chr(65 + i)}.  {opt}", key=f"opt_{idx}_{i}",
                         use_container_width=True):
                st.session_state.learn_answered = i
                st.session_state.learn_answers.append({
                    "question": q["question"],
                    "chosen": i,
                    "correct": i == q["answer_index"],
                    "answer_index": q["answer_index"],
                    "options": q["options"],
                    "explanation": q.get("explanation", ""),
                    "source_ref": q.get("source_ref", ""),
                })
                rerun()
        _end_session_button()
    else:
        # Reveal: color every option, then show the explanation
        for i, opt in enumerate(q["options"]):
            cls = "learn-option-neutral"
            mark = ""
            if i == q["answer_index"]:
                cls, mark = "learn-option-correct", "✓ "
            elif i == answered:
                cls, mark = "learn-option-wrong", "✗ "
            st.markdown(
                f"<div class='learn-option {cls}'>{mark}{chr(65 + i)}.  {opt}</div>",
                unsafe_allow_html=True,
            )

        right = answered == q["answer_index"]
        verdict = ("<span class='badge badge-success'>CORRECT</span>" if right
                   else "<span class='badge badge-error'>NOT QUITE</span>")
        src = q.get("source_ref", "")
        src_html = (f"<div style='font-size:11px;color:#666;margin-top:6px;'>"
                    f"📍 {src}</div>" if src else "")
        st.markdown(
            f"<div class='learn-explanation'>{verdict}<br>"
            f"{q.get('explanation', '')}{src_html}</div>",
            unsafe_allow_html=True,
        )

        last = idx + 1 >= len(items)
        if st.button("🏁 See results" if last else "Next →",
                     type="primary", key=f"next_{idx}"):
            st.session_state.learn_idx += 1
            st.session_state.learn_answered = None
            if last:
                st.session_state.learn_phase = "done"
            rerun()


# ── Flashcards ────────────────────────────────────────────────────────────────

def _render_flashcards():
    items = st.session_state.learn_items
    idx = st.session_state.learn_idx

    # When the main deck is exhausted, missed cards come back around
    if idx >= len(items):
        if st.session_state.learn_again:
            items.extend(st.session_state.learn_again)
            st.session_state.learn_again = []
        else:
            st.session_state.learn_phase = "done"
            rerun()
            return

    card = items[idx]
    _progress_header("Card")

    if not st.session_state.learn_flipped:
        _question_card(card["front"])
        if card.get("hint"):
            with st.expander("💡 Hint"):
                st.markdown(f"<span style='font-size:13px;color:#aaa;'>"
                            f"{card['hint']}</span>", unsafe_allow_html=True)
        if st.button("🔄 Flip card", type="primary", use_container_width=True):
            st.session_state.learn_flipped = True
            rerun()
        _end_session_button()
    else:
        _question_card(card["front"])
        src = card.get("source_ref", "")
        src_html = (f"<div style='font-size:11px;color:#666;margin-top:6px;'>"
                    f"📍 {src}</div>" if src else "")
        st.markdown(
            f"<div class='learn-flash-back'>{card['back']}{src_html}</div>",
            unsafe_allow_html=True,
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✗  Again later", use_container_width=True):
                st.session_state.learn_again.append(card)
                st.session_state.learn_answers.append(
                    {"correct": False, "source_ref": card.get("source_ref", "")})
                _next_card()
        with c2:
            if st.button("✓  Got it", type="primary", use_container_width=True):
                st.session_state.learn_answers.append(
                    {"correct": True, "source_ref": card.get("source_ref", "")})
                _next_card()


def _next_card():
    st.session_state.learn_idx += 1
    st.session_state.learn_flipped = False
    rerun()


# ── Interview ─────────────────────────────────────────────────────────────────

def _render_interview():
    items = st.session_state.learn_items
    idx = st.session_state.learn_idx
    if idx >= len(items):
        st.session_state.learn_phase = "done"
        rerun()
        return

    q = items[idx]
    _progress_header("Question")
    _question_card(q["question"], q.get("difficulty", ""))

    ev = st.session_state.learn_eval
    if ev is None:
        # 🎙️ Spoken answer: record → Whisper transcript lands in the text box
        # below (still editable before submitting). This must run BEFORE the
        # text_area is instantiated so the transcript can be injected into it.
        if hasattr(st, "audio_input"):
            audio = st.audio_input(
                "🎙️ Speak your answer — it will appear below as text",
                key=f"mic_{idx}",
            )
            if audio is not None:
                wav = audio.getvalue()
                digest = f"{len(wav)}-{hash(wav[:512])}"
                if st.session_state.get(f"mic_done_{idx}") != digest:
                    with st.spinner("Transcribing your answer…"):
                        transcript = transcribe_audio(wav)
                    st.session_state[f"mic_done_{idx}"] = digest
                    if transcript:
                        prior = st.session_state.get(f"ans_{idx}", "")
                        st.session_state[f"ans_{idx}"] = (
                            f"{prior} {transcript}".strip())
                    else:
                        st.warning("Couldn't transcribe that recording — "
                                   "try again or type your answer.")

        answer = st.text_area("Your answer", key=f"ans_{idx}", height=140,
                              placeholder="Speak with the mic above, or type "
                                          "as you would answer in the interview…")
        if st.button("Submit answer", type="primary"):
            with st.spinner("Coach is reading your answer…"):
                result = _run_async(evaluate_interview_answer(q, answer))
            st.session_state.learn_eval = result
            st.session_state.learn_answers.append({
                "question": q["question"],
                "answer": answer,
                "score": result["score"],
                "correct": result["score"] >= 6,
                "source_ref": q.get("source_ref", ""),
            })
            rerun()
        _end_session_button()
    else:
        score = ev["score"]
        badge = ("badge-success" if score >= 8
                 else "badge-info" if score >= 5 else "badge-error")
        covered = "".join(f"<li>✓ {c}</li>" for c in ev.get("covered_points", []))
        missed = "".join(f"<li>✗ {m}</li>" for m in ev.get("missed_points", []))
        st.markdown(
            f"<div class='learn-explanation learn-coach'>"
            f"<span class='badge {badge}'>SCORE {score}/10</span> "
            f"<span style='font-size:11px;color:#888;'>🎤 coach feedback</span><br>"
            f"{ev['feedback']}"
            f"<ul class='learn-points'>{covered}{missed}</ul></div>",
            unsafe_allow_html=True,
        )
        if q.get("model_answer"):
            with st.expander("📖 Model answer"):
                st.markdown(q["model_answer"])

        last = idx + 1 >= len(items)
        if st.button("🏁 See results" if last else "Next question →",
                     type="primary"):
            st.session_state.learn_idx += 1
            st.session_state.learn_eval = None
            if last:
                st.session_state.learn_phase = "done"
            rerun()


# ── Review mistakes (quiz) ────────────────────────────────────────────────────

def _render_review():
    wrong = [a for a in st.session_state.learn_answers
             if not a.get("correct") and a.get("options")]
    st.markdown("<p class='learn-section-label'>REVIEW YOUR MISTAKES</p>",
                unsafe_allow_html=True)
    if not wrong:
        st.markdown("Nothing to review — perfect round! 🎉")
    for a in wrong:
        correct_opt = a["options"][a["answer_index"]]
        chosen_opt = a["options"][a["chosen"]]
        src = a.get("source_ref", "")
        src_html = (f"<div style='font-size:11px;color:#666;margin-top:6px;'>"
                    f"📍 {src}</div>" if src else "")
        st.markdown(
            f"<div class='learn-question' style='margin-bottom:4px;'>{a['question']}</div>"
            f"<div class='learn-option learn-option-wrong'>✗ Your answer: {chosen_opt}</div>"
            f"<div class='learn-option learn-option-correct'>✓ Correct: {correct_opt}</div>"
            f"<div class='learn-explanation'>{a.get('explanation', '')}{src_html}</div>",
            unsafe_allow_html=True,
        )
    if st.button("← Back to results", type="primary"):
        st.session_state.learn_phase = "done"
        rerun()


# ── Results ───────────────────────────────────────────────────────────────────

def _render_done():
    answers = st.session_state.learn_answers
    mode = st.session_state.learn_mode
    total = len(answers)

    st.markdown("<p class='learn-section-label'>SESSION RESULTS</p>",
                unsafe_allow_html=True)

    if total == 0:
        st.markdown("Session ended before any answers.")
    elif mode == "interview":
        avg = sum(a.get("score", 0) for a in answers) / total
        label, badge = grade_label(avg * 10)
        st.markdown(
            f"<div class='learn-result'>"
            f"<div class='learn-result-score'>{avg:.1f}<span>/10 avg</span></div>"
            f"<span class='badge {badge}'>{label}</span>"
            f"<div style='font-size:12px;color:#888;margin-top:8px;'>"
            f"{total} question(s) · {st.session_state.learn_doc_title}</div></div>",
            unsafe_allow_html=True,
        )
    else:
        right = sum(1 for a in answers if a.get("correct"))
        pct = right / total * 100
        label, badge = grade_label(pct)
        st.markdown(
            f"<div class='learn-result'>"
            f"<div class='learn-result-score'>{right}<span>/{total}</span></div>"
            f"<span class='badge {badge}'>{label}</span>"
            f"<div style='font-size:12px;color:#888;margin-top:8px;'>"
            f"{pct:.0f}% · {st.session_state.learn_doc_title}</div></div>",
            unsafe_allow_html=True,
        )

    weak = weak_areas(answers)
    if weak:
        rows = "".join(f"<li>📍 {ref} — missed {n}</li>" for ref, n in weak)
        st.markdown(
            f"<div class='learn-explanation'><strong>Focus next on:</strong>"
            f"<ul class='learn-points'>{rows}</ul></div>",
            unsafe_allow_html=True,
        )

    c1, c2, c3 = st.columns(3)
    with c1:
        wrong_quiz = [a for a in answers if not a.get("correct") and a.get("options")]
        if mode == "quiz" and wrong_quiz:
            if st.button("📝 Review mistakes", use_container_width=True):
                st.session_state.learn_phase = "review"
                rerun()
    with c2:
        if st.button("🔁 New session", use_container_width=True, type="primary"):
            # Same settings; weak areas become the focus for the next round
            if weak and not st.session_state.learn_settings.get("focus"):
                st.session_state.learn_settings["focus"] = ", ".join(
                    ref for ref, _ in weak)
            _start_session(
                st.session_state.get("source_doc_path"),
                st.session_state.get("learn_paste", ""),
            )
    with c3:
        if st.button("🏠 Change mode", use_container_width=True):
            st.session_state.learn_phase = "setup"
            rerun()
