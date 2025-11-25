# PPT Copilot backend

Minimal FastAPI service that proxies chat requests to Gemini 3.

## Prerequisites
- Python 3.10+
- Environment variable `GEMINI_API_KEY` (can be loaded from a `.env` file)
- Optional: `ALLOWED_ORIGINS` comma-separated (defaults to `https://localhost:5173`)

## Setup
```bash
python -m venv .venv
. .venv/Scripts/activate  # Windows PowerShell: .venv\\Scripts\\Activate.ps1
pip install -r backend/requirements.txt
# Optional: create backend/.env with GEMINI_API_KEY=...
```

## Run locally
```bash
uvicorn backend.main:app --reload --port 8000
```

## Smoke tests
- Healthcheck:
```bash
curl http://localhost:8000/health
```
- Chat (expects backend + valid GEMINI_API_KEY):
```bash
curl -X POST http://localhost:8000/api/chat \
  -H "Content-Type: application/json" \
  -d '{ "messages": [ { "role": "user", "content": "CONTEXTE SLIDE (selection PowerPoint) :\nHello\n\nINSTRUCTION UTILISATEUR :\nRaccourcis ce texte en 2 bullets." } ] }'
```

## Endpoints
- `GET /health` -> `{ "status": "ok" }`
- `POST /api/chat` with body:
```json
{ "messages": [ { "role": "user", "content": "..." } ] }
```
Returns `{ "assistant_text": "..." }`.

## Notes
- CORS is open (`*`) for development; restrict before deploying.
- The system prompt is defined in `backend/system_instruction.py`.
