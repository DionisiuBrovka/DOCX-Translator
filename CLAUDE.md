# CLAUDE.md

## Project Overview

DOCX Translator is a Python/Flask web application that translates Word documents (.docx) between languages (default: Russian → English) using the Groq LLM API while preserving all document formatting. The core technique is **run-level translation** — individual formatting fragments within paragraphs are translated as structured JSON arrays, keeping bold/italic/underline/font attributes intact.

## Tech Stack

- **Backend**: Python 3, Flask >= 3.0
- **Document processing**: python-docx >= 1.1
- **Translation API**: Groq >= 0.9 (LLM-based translation)
- **Frontend**: Vanilla HTML/CSS/JS (single-page in `templates/index.html`)
- **No build step** — pure Python, no compilation needed

## Project Structure

```
├── app.py                 # Entire backend: TranslationEngine, DocxTranslator, Flask routes
├── templates/
│   └── index.html         # Complete frontend (HTML + CSS + JS, dark theme)
├── requirements.txt       # 3 dependencies: flask, python-docx, groq
├── .env.example           # Environment variable template
├── README.md              # Documentation (in Russian)
├── uploads/               # Temporary upload directory (auto-created, files cleaned after processing)
└── outputs/               # Translated documents stored here
```

## Architecture

All backend logic lives in `app.py` (~450 lines), organized into three sections:

1. **`TranslationEngine`** (lines 42–180) — Handles Groq API communication
   - `translate_runs()`: Entry point; dispatches to simple or structured translation
   - `_translate_simple()`: Single text fragment translation with whitespace preservation
   - `_translate_structured()`: Multi-fragment JSON array translation to maintain formatting boundaries
   - Fallback: if JSON parsing fails, translates full paragraph as one block

2. **`DocxTranslator`** (lines 187–316) — Traverses .docx structure
   - Processes body paragraphs, tables (including nested), headers, and footers
   - Skips non-Cyrillic text (detects via Unicode range `\u0400`–`\u04FF`)
   - Tracks progress via callback for real-time UI updates

3. **Flask routes** (lines 322–442):
   - `GET /` — Serves the web UI
   - `POST /translate` — Upload and translate a document (synchronous)
   - `GET /progress/<task_id>` — Poll translation progress
   - `GET /download/<task_id>` — Download translated document

## Setup & Running

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Set API key (or enter via web UI)
export GROQ_API_KEY="gsk_..."
python app.py
# Open http://localhost:5000
```

## Environment Variables

| Variable       | Default                      | Purpose                  |
|----------------|------------------------------|--------------------------|
| `GROQ_API_KEY` | (none)                       | Groq API key (required)  |
| `GROQ_MODEL`   | `llama-3.3-70b-versatile`   | LLM model for translation|
| `PORT`         | `5000`                       | Web server port          |
| `FLASK_DEBUG`  | `0`                          | Flask debug mode         |
| `SECRET_KEY`   | hardcoded default            | Flask session secret     |

## Key Design Decisions

- **Run-level translation**: Preserves inline formatting by translating each `run` (formatting fragment) independently via JSON arrays sent to the LLM.
- **Synchronous processing**: Translation blocks the request thread. Not suitable for production without a task queue (Celery/RQ).
- **In-memory task storage**: Progress and results stored in a Python dict. Lost on restart.
- **Cyrillic detection**: Paragraphs without Cyrillic characters are skipped, assuming they're already in the target language.
- **Temperature tuning**: 0.3 for simple translations, 0.2 for structured (lower = more deterministic).

## Testing

No test suite exists. There are no test files, test configuration, or CI/CD pipelines.

When adding tests:
- Use `pytest` as the test framework
- Mock the `Groq` client to avoid real API calls
- Key areas to test: `TranslationEngine` JSON parsing/fallback logic, `DocxTranslator` paragraph/table traversal, Flask route validation

## Common Tasks

**Adding a new route**: Add to the Flask routes section at the bottom of `app.py` (after line 316).

**Changing translation behavior**: Modify `TranslationEngine` methods. The system prompts for the LLM are inline in `_translate_simple()` and `_translate_structured()`.

**Modifying the UI**: Edit `templates/index.html` — all styles, markup, and JavaScript are in this single file.

**Adding a new dependency**: Add to `requirements.txt` and run `pip install -r requirements.txt`.

## Code Conventions

- Comments and docstrings are in Russian (project originated in Russian-speaking context)
- Type hints are used sparingly (e.g., `list[str]`, `Optional`)
- Logging via Python's `logging` module (`logger = logging.getLogger(__name__)`)
- Paths use `pathlib.Path` rather than string manipulation
- No linter or formatter is configured — follow existing style (PEP 8 general adherence)

## Known Limitations

- Text inside embedded objects (SmartArt, charts) is not translated
- Text on images is not translated (would need OCR)
- Very long paragraphs may exceed LLM token limits
- Synchronous processing blocks the server during translation
- No persistent storage — task state is lost on server restart
- No authentication or rate limiting
