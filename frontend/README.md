# PPT Copilot taskpane (frontend)

React + Vite taskpane that talks to the FastAPI backend and uses Office JS to read/write the selected PowerPoint shape.

## Prerequisites
- Node.js 18+
- PowerPoint with Web Add-ins enabled
- Backend running (default `http://localhost:8000`)

## Setup
```bash
cd frontend
npm install
cp .env.example .env.local  # adjust API base URL if needed
```

## Run
```bash
npm run dev
```
The dev server runs on `https://localhost:5173` (self-signed cert via Vite basic-ssl).

## Notes
- Uses `office-js` and Office globals (`PowerPoint.run`) for selection read/write.
- Buttons: send chat, apply last assistant response to selection, Undo IA per shape.
- Message format matches backend expectations (CONTEXTE SLIDE + INSTRUCTION UTILISATEUR).

## Side-load in PowerPoint (desktop)
1) Ensure the backend is running on `http://localhost:8000`.
2) Start the dev server: `npm run dev` (accept the self-signed cert in the browser once).
3) In PowerPoint, go to **Insert > My Add-ins > Upload My Add-in** (or **Shared Folder** depending on your setup).
4) Select `manifest.xml` at the repo root; the taskpane points to `https://localhost:5173/index.html`.
5) Open the taskpane and keep the dev server running while testing.

## Quick e2e test
1) In PowerPoint, type some text in a shape and select it.
2) In the taskpane, enter a prompt (ex: "Raccourcis ce texte en 3 bullets.").
3) Click **Envoyer** → receive the assistant response.
4) Click **Appliquer à la sélection** → shape text is replaced, undo entry saved.
5) Click **Undo IA** → previous text restored. Handle cases: no selection, no undo entry, backend down.
