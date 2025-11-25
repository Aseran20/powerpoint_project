# Guide développeur – Add-in "PPT Copilot M&A"

Ce document décrit ce qu’il faut coder pour la V1 de l’add-in PowerPoint connecté à Gemini 3.

---

## 1. Objectif du projet

Construire un add-in PowerPoint qui :
- Affiche un panneau latéral de chat
- Utilise Gemini 3 pour générer / réécrire du texte M&A
- Utilise automatiquement le texte de la shape sélectionnée comme contexte
- Permet d’appliquer le dernier message de l’IA dans la shape sélectionnée
- Permet d’annuler la dernière modification IA par shape (Undo IA)

Pas de boutons "raccourcir", "bullets", etc. dans la V1. L’utilisateur travaille en mode chat libre (comme Copilot / Codex), plus un bouton d’application et un Undo.

---

## 2. Architecture d’ensemble

Deux blocs principaux :

- **Add-in PowerPoint**
  - Panneau HTML/JS (React recommandé)
  - Office JS pour lire / modifier le texte de la shape sélectionnée
  - Historique du chat maintenu côté client
  - Appels HTTP au backend `/api/chat`

- **Backend Python**
  - FastAPI (ou équivalent)
  - SDK Gemini 3 : `google-genai`
  - Endpoint `/api/chat` qui applique la `system_instruction` M&A
  - Appel du modèle `gemini-3-pro-preview`

Pas de base de données obligatoire pour la V1 (historique uniquement côté client).

---

## 3. Backend Gemini 3

### 3.1. Prérequis

- Python 3.10+
- Paquets :
  - `fastapi`
  - `uvicorn`
  - `pydantic`
  - `google-genai`
- Variable d’environnement : `GEMINI_API_KEY`

Installation :

```bash
pip install fastapi uvicorn pydantic google-genai
```

### 3.2. Initialisation du client

```python
import os
from google import genai

api_key = os.environ["GEMINI_API_KEY"]
client = genai.Client(api_key=api_key)

MODEL_NAME = "gemini-3-pro-preview"
```

### 3.3. System instruction

Créer un fichier `system_instruction.py` avec une constante `SYSTEM_PROMPT` (voir annexe complète en bas du document).

```python
SYSTEM_PROMPT = """
...[coller la system_instruction complète ici]...
"""
```

### 3.4. Schémas de requête / réponse API

`POST /api/chat`

**Body JSON attendu :**

```json
{
  "messages": [
    { "role": "user", "content": "..." },
    { "role": "assistant", "content": "..." }
  ]
}
```

- `messages` = historique complet du chat côté panneau, dans l’ordre
- `content` = texte brut
- Les messages `user` doivent déjà être au format :

```text
CONTEXTE SLIDE (sélection PowerPoint) :
{texte_selection}

INSTRUCTION UTILISATEUR :
{prompt_utilisateur}
```

**Réponse JSON :**

```json
{
  "assistant_text": "...",
  "usage": {
    "input_tokens": 123,
    "output_tokens": 456
  }
}
```

`usage` est optionnel (logging / métriques).

### 3.5. Modèles Pydantic

```python
from pydantic import BaseModel
from typing import List, Literal, Optional

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
```

### 3.6. Endpoint `/api/chat`

```python
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from google import genai
from system_instruction import SYSTEM_PROMPT
from models import ChatRequest, ChatResponse
import os

api_key = os.environ["GEMINI_API_KEY"]
client = genai.Client(api_key=api_key)
MODEL_NAME = "gemini-3-pro-preview"

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # à restreindre plus tard
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/chat", response_model=ChatResponse)
def chat(req: ChatRequest):
    if not req.messages:
        raise HTTPException(status_code=400, detail="messages is required")

    contents = []
    for msg in req.messages:
        contents.append({
            "role": msg.role,
            "parts": [{"text": msg.content}],
        })

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
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    assistant_text = response.text if hasattr(response, "text") else ""

    input_tokens = None
    output_tokens = None

    return ChatResponse(
        assistant_text=assistant_text,
        input_tokens=input_tokens,
        output_tokens=output_tokens,
    )
```

### 3.7. Healthcheck

```python
@app.get("/health")
def health():
    return {"status": "ok"}
```

### 3.8. Lancement local

```bash
uvicorn main:app --reload --port 8000
```

---

## 4. Add-in PowerPoint

Objectif : un taskpane React minimal qui :
- Affiche un chat
- Connaît la shape sélectionnée
- Envoie `messages` au backend
- Permet d’appliquer la dernière réponse IA à la shape
- Permet un Undo IA par shape

### 4.1. Prérequis

- Node.js récent
- PowerPoint avec support Web Add-ins
- Office JS

### 4.2. Modèle de données côté client

```ts
type Role = "user" | "assistant"

interface ChatMessage {
  id: string
  role: Role
  content: string
  createdAt: number
}

interface UndoEntry {
  shapeId: string
  previousText: string
}

interface UiState {
  messages: ChatMessage[]
  lastAssistantMessageId: string | null
  lastUndoByShape: Record<string, UndoEntry | undefined>
}
```

### 4.3. Fonctions Office JS principales

Lire le texte de la shape sélectionnée :

```ts
async function getSelectedShapeText(): Promise<{ shapeId: string | null; text: string }> {
  return PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes()
    selectedShapes.load("items")
    await context.sync()

    if (selectedShapes.items.length === 0) {
      return { shapeId: null, text: "" }
    }

    const shape = selectedShapes.items[0]
    shape.load("id, textFrame/textRange/text")
    await context.sync()

    const shapeId = shape.id
    const text = shape.textFrame.textRange.text || ""

    return { shapeId, text }
  })
}
```

Remplacer le texte de la shape sélectionnée :

```ts
async function setSelectedShapeText(newText: string): Promise<{ shapeId: string | null }> {
  return PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes()
    selectedShapes.load("items")
    await context.sync()

    if (selectedShapes.items.length === 0) {
      return { shapeId: null }
    }

    const shape = selectedShapes.items[0]
    shape.load("id, textFrame/textRange")
    await context.sync()

    shape.textFrame.textRange.text = newText
    await context.sync()

    return { shapeId: shape.id }
  })
}
```

(L’API exacte peut varier légèrement selon la version Office JS, mais l’idée est : `getSelectedShapes()` → premier shape → `id` + `textFrame.textRange.text`.)

### 4.4. Client HTTP `/api/chat`

```ts
import type { ChatMessage } from "./types"

const API_BASE_URL = "http://localhost:8000" // à adapter

export async function sendChat(messages: ChatMessage[]) {
  const body = {
    messages: messages.map(m => ({
      role: m.role,
      content: m.content,
    })),
  }

  const res = await fetch(`${API_BASE_URL}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  })

  if (!res.ok) {
    throw new Error(`Backend error: ${res.status}`)
  }

  return res.json() as Promise<{ assistant_text: string }>
}
```

### 4.5. Construction du message `user`

Au click sur "Envoyer" :

1. Lire le texte de la sélection :

```ts
const { shapeId, text: selectionText } = await getSelectedShapeText()
```

2. Construire `content` :

```ts
const content = [
  "CONTEXTE SLIDE (sélection PowerPoint) :",
  selectionText || "",
  "",
  "INSTRUCTION UTILISATEUR :",
  userPrompt.trim(),
].join("\n")
```

3. Ajouter au state `messages` comme message `user`, puis envoyer tout l’historique au backend.

### 4.6. Réception de la réponse IA

Après `sendChat` :

```ts
const { assistant_text } = await sendChat([...messages, newUserMessage])

const newAssistantMessage: ChatMessage = {
  id: generateRandomId(),
  role: "assistant",
  content: assistant_text,
  createdAt: Date.now(),
}

setState(prev => ({
  ...prev,
  messages: [...prev.messages, newUserMessage, newAssistantMessage],
  lastAssistantMessageId: newAssistantMessage.id,
}))
```

### 4.7. Bouton "Appliquer à la sélection"

Logique :

1. Prendre le dernier message `assistant`
2. Lire la shape sélectionnée et son texte actuel
3. Sauvegarder ce texte dans `lastUndoByShape[shapeId]`
4. Remplacer le texte de la shape par `assistant.content`

Exemple :

```ts
async function applyLastAssistantMessage(state: UiState, setState: (s: UiState) => void) {
  const lastAssistant = state.messages.filter(m => m.role === "assistant").slice(-1)[0]
  if (!lastAssistant) return

  const { shapeId, text: currentText } = await getSelectedShapeText()
  if (!shapeId) return

  const undoEntry: UndoEntry = {
    shapeId,
    previousText: currentText,
  }

  setState(prev => ({
    ...prev,
    lastUndoByShape: {
      ...prev.lastUndoByShape,
      [shapeId]: undoEntry,
    },
  }))

  await setSelectedShapeText(lastAssistant.content)
}
```

### 4.8. Bouton "Undo IA"

Logique :

1. Prendre la shape sélectionnée
2. Vérifier s’il existe un `UndoEntry` pour ce `shapeId`
3. Si oui, réécrire `previousText` dans la shape et nettoyer l’entrée

Exemple :

```ts
async function undoIaForSelection(state: UiState, setState: (s: UiState) => void) {
  const { shapeId } = await getSelectedShapeText()
  if (!shapeId) return

  const undoEntry = state.lastUndoByShape[shapeId]
  if (!undoEntry) return

  await setSelectedShapeText(undoEntry.previousText)

  setState(prev => {
    const clone = { ...prev.lastUndoByShape }
    delete clone[shapeId]
    return { ...prev, lastUndoByShape: clone }
  })
}
```

---

## 5. Format des messages envoyés au modèle

Règles à respecter côté add-in :

- Ne pas envoyer de JSON ou de méta dans le `content`, uniquement du texte brut.
- Toujours respecter le pattern suivant dans les messages `user` :

```text
CONTEXTE SLIDE (sélection PowerPoint) :
{texte_de_la_shape}

INSTRUCTION UTILISATEUR :
{prompt_que_l_utilisateur_a_tapé}
```

- Le `system_instruction` se charge du ton M&A, des bullets `- `, de la numérotation `1.`, du gras `**...**`, et du fait de ne pas ajouter de phrases du type "Voici votre texte".

---

## 6. Scénario de test end-to-end

1. L’utilisateur tape un bloc de texte long dans une shape.
2. Il sélectionne cette shape.
3. Il ouvre le panneau de l’addin, écrit par exemple :
   "Raccourcis ce texte en 4 bullet points pour un IM"
4. Le front :
   - lit `selectionText`
   - construit le message `user` avec CONTEXTE + INSTRUCTION
   - envoie l’historique complet au backend
5. Le backend :
   - appelle Gemini 3 avec `SYSTEM_PROMPT` + `contents`
   - renvoie `assistant_text` (par ex. une liste de bullets `- ...`)
6. Le panneau affiche la réponse IA.
7. L’utilisateur clique "Appliquer" :
   - le front sauvegarde `previousText` pour cette shape
   - remplace le texte de la shape par la réponse IA
8. L’utilisateur clique "Undo IA" :
   - le front réinjecte `previousText` dans la shape.

---

## 7. Checklist V1

Backend :
- Endpoint `/health` opérationnel
- Endpoint `POST /api/chat` fonctionnel avec Gemini 3
- `SYSTEM_PROMPT` correctement configuré en `system_instruction`
- CORS configuré pour autoriser le domaine de l’addin

Add-in PowerPoint :
- Panneau latéral affiché dans PowerPoint
- Saisie du prompt + affichage de l’historique du chat
- Lecture du texte de la shape sélectionnée
- Envoi des `messages` au backend au bon format
- Affichage de la réponse IA
- Bouton "Appliquer à la sélection" opérationnel
- Undo IA fonctionnel par shape
- Messages d’erreur minimaux (pas de shape sélectionnée, backend indisponible, etc.)

---

## 8. Annexe – System instruction à utiliser

À mettre côté backend dans `SYSTEM_PROMPT` :

```text
You are a slide-writing assistant integrated into Microsoft PowerPoint.

Your environment and role
- You are used inside a PowerPoint add-in on professional finance / M&A presentations (teasers, information memorandums, pitchbooks, board decks, investor updates).
- Your main job is to help the user write, rewrite, shorten, clarify, or translate the text that will appear inside a single selected text box on a slide.
- Typical content includes: company descriptions, key investment highlights, market overviews, strategic rationales, process descriptions, financial summaries, and transaction terms.

Input structure
- You receive conversational messages from the user.
- Some user messages may include two explicit parts:
  - A block starting with "CONTEXTE SLIDE (sélection PowerPoint) :" – this is the current text in the selected shape on the slide. This block may be empty.
  - A block starting with "INSTRUCTION UTILISATEUR :" – this describes what the user wants you to do.
- When both are present:
  - Treat the CONTEXTE SLIDE as the base text you should transform, improve, shorten, translate, or structure.
  - Treat the INSTRUCTION UTILISATEUR as the precise task description and style guidelines you must follow.
- When there is no CONTEXTE SLIDE, you may be asked to generate text from scratch (for example: write bullets given only a high-level instruction).

Tone and style
- Always write in a professional, concise finance / M&A tone, suitable for investment banks, private equity funds, and corporate finance teams.
- Preserve factual content (numbers, company names, places, dates, transaction terms) unless the user explicitly asks you to change or anonymize them.
- Prefer clarity and impact over marketing fluff. Avoid buzzwords and vague language.
- Default to short, slide-ready formulations: clear sentences, no unnecessary introductions, no apologies, no meta-commentary.
- When the user asks for a shorter or more concise version, reduce length while keeping all key ideas.

Language
- Use the language requested by the user.
- If no explicit language is requested, use the same main language as the user’s latest message:
  - French if the user writes in French.
  - English if the user writes in English.
- For translation tasks, preserve the meaning and tone, and adapt to standard corporate / M&A wording in the target language.

Formatting rules (very important)
- Your output will be pasted directly into a single PowerPoint text box. It must be immediately usable.
- Unless the user explicitly asks for something else:
  - Do NOT add titles, headers, or section labels.
  - Do NOT add explanations like "Here is your text" or "I have rewritten".
  - Do NOT include markdown headings (no lines starting with "# ").
- Bullet points
  - When the user asks for bullet points or a list, return one bullet per line.
  - Start each bullet with "- " (dash + space) and then the text of the bullet.
  - Do not add extra blank lines between bullets unless the user asks for them.
- Numbered lists
  - If the user asks explicitly for a numbered list, use the format:
    - "1. ..."
    - "2. ..."
    - etc., one item per line.
- Bold / emphasis
  - When the user asks to highlight or emphasize elements, use Markdown-style bold with **double asterisks** around the text to be bolded, for example: La société **Bubbles Crèches**.
  - Do not use any other markdown formatting (no italics, no links) unless the user explicitly requires it.
- Tables
  - Avoid tables unless the user explicitly asks for a table-like format.
  - If a table is required, use a simple text representation (for example pipe-separated), suitable for manual post-processing.

Content constraints
- Stay strictly within the user-provided information when rewriting, shortening, or translating. Do not invent financials, company names, or transaction details.
- If the request could lead to fabricating numbers or facts, keep the wording generic or rephrase what is already there instead of inventing data.
- When summarizing or shortening, keep:
  - Key quantitative metrics (revenue, EBITDA, growth, margins, multiples)
  - Strategic differentiators
  - Major milestones and deal terms if present
- Do not add disclaimers unless explicitly requested.

Behavior in conversation
- The add-in behaves like a chat: the user may refine or iterate on your previous answer.
- You should remember the conversation context and the user’s preferences within the current conversation, but always prioritize the latest instruction.
- If an instruction conflicts with the default rules above, follow the instruction as long as it is explicit and unambiguous.
- If the user asks for style-specific output (for example: "style IM", "style board", "style teaser", "more marketing", "more direct"), adapt your tone accordingly while staying professional.

Error handling and uncertainty
- If the user’s instruction is unclear but you can reasonably infer what is meant, choose the most slide-appropriate interpretation and proceed.
- If you genuinely lack information needed for a factual element (for example a missing number) and the user expects you to keep content factual, avoid inventing and write around the missing detail.

Overall objective
- Your primary objective is to produce slide-ready text that the user can paste directly into the selected PowerPoint text box, with minimal or no manual editing.
- Always favor clarity, concision, and professional M&A style, while respecting the user’s specific instructions and the constraints above.
```

