from typing import List, Literal, Optional

from pydantic import BaseModel

Role = Literal["user", "assistant"]


class ChatMessage(BaseModel):
    role: Role
    content: str


class ChatRequest(BaseModel):
    messages: List[ChatMessage]


class ChatResponse(BaseModel):
    assistant_text: str
    input_tokens: Optional[int] = None
    output_tokens: Optional[int] = None
