# ⚡ DocuForge

**A multi-agent AI system that reformats and enhances business documents — preserving your exact theme, fonts, and colors while fixing, restructuring, and generating new content.**

DocuForge takes a `.docx`, `.pptx`, or `.pdf`, understands a natural-language instruction
("apply this template", "fix spellings and build a RACI matrix", "follow the design of
slide 4 everywhere"), and produces a polished output that looks hand-crafted — no prompt
engineering required.

> Built on top of [JoshuaC215/agent-service-toolkit](https://github.com/JoshuaC215/agent-service-toolkit)
> (LangGraph + FastAPI + Streamlit), extended with a document pipeline, Supabase auth,
> a self-correcting critic, and persistent memory.
> The original toolkit docs are preserved in [`docs/UPSTREAM_TOOLKIT.md`](docs/UPSTREAM_TOOLKIT.md).

---

## ✨ What it does

- **Template-mode editing** — opens your source file and edits it *in place*, so the
  original theme, fonts, colors, cover page, headers/footers, and existing tables are
  **preserved by construction**. No "rebuild from scratch" that flattens your branding.
- **RACI matrix generation** — analyzes a process document and builds a color-coded
  Responsible/Accountable/Consulted/Informed matrix (activities × stakeholders), grounded
  in the document's actual content.
- **Process-flow generation** — produces a styled vertical flow chart (stakeholder → step,
  with arrows) from the described end-to-end process.
- **Spelling & consistency fixes** — run-level corrections that keep all original formatting.
- **PPTX theme application** — apply one slide's design (fonts, colors, table styles) across
  an entire deck; fills in missing agenda/index slides.
- **PDF extraction** — pull content out of PDFs and reformat into `.docx`.
- **Web research with citations**, **persistent semantic memory**, and a **critic agent**
  that reviews every output before it's returned.

## 🧠 How it works

```
User message
     │
     ▼
Memory Agent ──► retrieves relevant past context (FAISS, optional)
     │
     ▼
Intent Router ──► doc_task_docx | doc_task_pptx | doc_task_pdf | research | chat | code
     │
     ├─ document ─► extract ─► (LLM build / enhance) ─► template-mode builder ─► validate
     ├─ research ─► Tavily search + cite
     └─ chat     ─► direct LLM response
     │
     ▼
Critic Agent ──► reviews draft, retries on failure (budgeted)
     │
     ▼
Final response  +  downloadable document
```

A shared `AgentState` (LangGraph `TypedDict`) flows through every node. Document tasks run
extraction and style-parsing in parallel, then build, then validate — with a self-correction
loop capped by a retry budget.

## 🛠️ Tech stack

| Layer | Tool |
|-------|------|
| Agent framework | LangGraph |
| LLMs | Groq (Llama 3.3 70B / Llama 4 Scout / Qwen QwQ) with Google Gemini fallback |
| Documents | python-docx, python-pptx, mammoth, pypdf, lxml |
| Web search | Tavily |
| Memory | FAISS + FastEmbed (gracefully disabled if unavailable) |
| Auth & DB | Supabase |
| UI | Streamlit |
| API | FastAPI |

## 🚀 Getting started

### 1. Prerequisites
- Python 3.10+
- A [Groq API key](https://console.groq.com) (free tier works)

### 2. Install
```bash
git clone https://github.com/Akshay-More-007/docuforge.git
cd docuforge
pip install -e .                          # base toolkit deps (pyproject.toml)
pip install -r requirements-docuforge.txt # DocuForge extras
```

### 3. Configure
Create a `.env` file in the project root:
```bash
# LLMs (at least one required)
GROQ_API_KEY=gsk_...
GOOGLE_API_KEY=AIza...          # optional fallback

# Web research (optional)
TAVILY_API_KEY=tvly-...

# Auth / storage (optional — required for the login flow)
SUPABASE_URL=https://xxxx.supabase.co
SUPABASE_ANON_KEY=eyJ...

# Model overrides (optional)
GROQ_MODEL_FAST=llama-3.3-70b-versatile
GROQ_MODEL_LONG=llama-4-scout-17b-16e-instruct
GROQ_MODEL_REASON=qwen-qwq-32b
GEMINI_MODEL=gemini-2.5-flash
```

### 4. Run
```bash
streamlit run src/app/main.py
```
Open http://localhost:8501, upload a document in the sidebar, and describe what you want.

## 📁 Project structure

```
src/
├── agents/        # intent_router, document_agent, research_agent,
│                  # critic_agent, memory_agent, fallback_agent
├── document/      # extractor, style_parser, builder (docx),
│                  # docx_template_builder (in-place + RACI/flow),
│                  # ppt_builder (template-mode), validator
├── graph/         # state, nodes, edges, graph_builder
├── llm/           # router + Groq / Gemini clients
├── memory/        # faiss_store, chat_memory
├── auth/          # Supabase auth, session, guards
└── app/           # Streamlit UI (pages, components, styles)
```

## 📝 Example prompts

- *"Fix all spelling errors, build a RACI matrix mapping each process area to the responsible
  teams, and add an end-to-end process flow chart — keep the exact formatting."*
- *"Follow the design theme of slide 4 and apply it everywhere; complete the missing index slide."*
- *"Convert this PDF into a clean Word document with proper headings."*

## 🙏 Credits

DocuForge is built on the excellent
[agent-service-toolkit](https://github.com/JoshuaC215/agent-service-toolkit) by Joshua Carroll,
which provides the LangGraph + FastAPI + Streamlit foundation.

## License

MIT — see [LICENSE](LICENSE).
