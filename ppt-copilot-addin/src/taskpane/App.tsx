import React, { useMemo, useState } from "react";
import { v4 as uuid } from "uuid";
import { sendChat } from "./api";
import { getSelectedShapeText, setSelectedShapeFormattedText } from "./office";
import type { ChatMessage, UiState, UndoEntry } from "./types";

function friendlyError(err: unknown): string {
  if (err instanceof TypeError && err.message.toLowerCase().includes("fetch")) {
    return "Backend injoignable. Vérifie qu'il tourne bien en local.";
  }
  return err instanceof Error ? err.message : "Erreur inconnue";
}

function buildUserContent(selectionText: string, prompt: string) {
  return [
    "CONTEXTE SLIDE (selection PowerPoint) :",
    selectionText || "",
    "",
    "INSTRUCTION UTILISATEUR :",
    prompt.trim(),
  ].join("\n");
}

function lastAssistant(messages: ChatMessage[]): ChatMessage | undefined {
  const assistants = messages.filter((m) => m.role === "assistant");
  return assistants[assistants.length - 1];
}

export default function App() {
  const [state, setState] = useState<UiState>({
    messages: [],
    lastAssistantMessageId: null,
    lastUndoByShape: {},
  });
  const [input, setInput] = useState("");
  const [sending, setSending] = useState(false);
  const [applying, setApplying] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const assistantMessage = useMemo(() => lastAssistant(state.messages), [state.messages]);

  async function handleSend() {
    if (!input.trim()) return;
    setSending(true);
    setError(null);

    try {
      const { text: selectionText } = await getSelectedShapeText();
      const userMessage: ChatMessage = {
        id: uuid(),
        role: "user",
        content: buildUserContent(selectionText, input),
        createdAt: Date.now(),
      };

      const chatHistory = [...state.messages, userMessage];
      const { assistant_text } = await sendChat(chatHistory);

      const assistantMsg: ChatMessage = {
        id: uuid(),
        role: "assistant",
        content: assistant_text,
        createdAt: Date.now(),
      };

      setState((prev) => ({
        ...prev,
        messages: [...chatHistory, assistantMsg],
        lastAssistantMessageId: assistantMsg.id,
      }));
      setInput("");
    } catch (err) {
      setError(friendlyError(err));
    } finally {
      setSending(false);
    }
  }

  async function handleApply() {
    if (!assistantMessage) return;
    setApplying(true);
    setError(null);
    try {
      const { shapeId, text: currentText } = await getSelectedShapeText();
      if (!shapeId) {
        setError("Aucune shape sélectionnée.");
        return;
      }

      const undoEntry: UndoEntry = {
        shapeId,
        previousText: currentText,
      };

      setState((prev) => ({
        ...prev,
        lastUndoByShape: { ...prev.lastUndoByShape, [shapeId]: undoEntry },
      }));

      await setSelectedShapeFormattedText(assistantMessage.content);
    } catch (err) {
      setError(friendlyError(err));
    } finally {
      setApplying(false);
    }
  }

  async function handleUndo() {
    setError(null);
    try {
      const { shapeId } = await getSelectedShapeText();
      if (!shapeId) {
        setError("Aucune shape sélectionnée.");
        return;
      }

      const undoEntry = state.lastUndoByShape[shapeId];
      if (!undoEntry) {
        setError("Aucun Undo IA disponible pour cette shape.");
        return;
      }

      await setSelectedShapeText(undoEntry.previousText);
      setState((prev) => {
        const clone = { ...prev.lastUndoByShape };
        delete clone[shapeId];
        return { ...prev, lastUndoByShape: clone };
      });
    } catch (err) {
      setError(friendlyError(err));
    }
  }

  return (
    <div className="app">
      <div className="panel">
        <div className="panel-header">PPT Copilot</div>

        <div className="messages">
          {state.messages.length === 0 && (
            <div className="status">Sélectionne une shape puis envoie un prompt.</div>
          )}
          {state.messages.map((msg) => (
            <div key={msg.id} className={`message ${msg.role}`}>
              <small>{msg.role === "assistant" ? "Assistant" : "Utilisateur"}</small>
              <div style={{ whiteSpace: "pre-wrap" }}>{msg.content}</div>
            </div>
          ))}
        </div>

        <div className="input-area">
          {error && <div className="error">{error}</div>}

          <textarea
            placeholder="Décris ce que tu veux faire sur la sélection..."
            value={input}
            onChange={(e) => setInput(e.target.value)}
            disabled={sending}
          />

          <div className="actions">
            <button className="primary" onClick={handleSend} disabled={sending}>
              {sending ? "Envoi..." : "Envoyer"}
            </button>
            <button className="secondary" onClick={handleApply} disabled={!assistantMessage || applying}>
              {applying ? "Application..." : "Appliquer à la sélection"}
            </button>
            <button className="ghost" onClick={handleUndo}>
              Undo IA
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
