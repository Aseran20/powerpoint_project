import type { ChatMessage } from "./types";

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL ?? "https://localhost:8000";

export async function sendChat(messages: ChatMessage[]) {
  const body = {
    messages: messages.map((m) => ({
      role: m.role,
      content: m.content,
    })),
  };

  const res = await fetch(`${API_BASE_URL}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Backend error ${res.status}: ${text}`);
  }

  return res.json() as Promise<{ assistant_text: string }>;
}
