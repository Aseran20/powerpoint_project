import logging
import os
from pathlib import Path
from typing import List

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from google import genai
from dotenv import load_dotenv

from backend.models import ChatMessage, ChatRequest, ChatResponse
from backend.system_instruction import SYSTEM_PROMPT

MODEL_NAME = "gemini-3-pro-preview"

# Load .env both from current working dir and backend/.env (local dev)
load_dotenv()
backend_env = Path(__file__).parent / ".env"
if backend_env.exists():
    load_dotenv(backend_env, override=False)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("ppt-copilot")

allowed_origins_env = os.getenv("ALLOWED_ORIGINS", "https://localhost:5173")
ALLOWED_ORIGINS = [o.strip() for o in allowed_origins_env.split(",") if o.strip()]

app = FastAPI(title="PPT Copilot Backend")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

_client: genai.Client | None = None


def get_client() -> genai.Client:
    """Lazily build the Gemini client to surface a clear error if the API key is missing."""
    global _client
    if _client is None:
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            raise HTTPException(status_code=500, detail="GEMINI_API_KEY is not set")
        _client = genai.Client(api_key=api_key)
    return _client


def to_genai_contents(messages: List[ChatMessage]) -> list[dict]:
    """Map chat history to the format expected by google-genai."""
    contents = []
    for msg in messages:
        contents.append(
            {
                "role": msg.role,
                "parts": [{"text": msg.content}],
            }
        )
    return contents


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/chat", response_model=ChatResponse)
def chat(req: ChatRequest):
    if not req.messages:
        raise HTTPException(status_code=400, detail="messages is required")

    client = get_client()
    contents = to_genai_contents(req.messages)

    try:
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=contents,
            system_instruction={"parts": [{"text": SYSTEM_PROMPT}]},
            generation_config={
                "thinking_level": "high",
                "max_output_tokens": 512,
                "temperature": 1.0,
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"Gemini error: {exc}") from exc

    assistant_text = getattr(response, "text", None) or ""
    if not assistant_text:
        raise HTTPException(status_code=500, detail="Empty response from Gemini")

    logger.info(
        "chat completed",
        extra={
            "messages_count": len(req.messages),
            "assistant_chars": len(assistant_text),
        },
    )

    # Token usage is left optional; the SDK may expose it via response.candidates[0].usage_metadata.
    return ChatResponse(
        assistant_text=assistant_text,
        input_tokens=None,
        output_tokens=None,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
