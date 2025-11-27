# PPT Copilot taskpane (Yo Office template)

Official Office taskpane scaffold for PowerPoint with our React UI wired to the FastAPI backend.

## Run (local dev)
```bash
npm install
npm run dev-server  # https://localhost:3000
```
- Backend: set `GEMINI_API_KEY` and start FastAPI on `https://localhost:8000`.
- API URL: override with `DEV_BACKEND_URL` (dev) or `PROD_BACKEND_URL`/`BACKEND_URL` (build). Default: `https://localhost:8000`.
- CORS: backend default allows 3000/5173.

## Sideload
- Use `manifest.xml` at repo root (copied from this template) or `ppt-copilot-addin/manifest.xml`.
- In PowerPoint: Insert > My Add-ins > Upload My Add-in, choose the manifest.
- Command bar button opens the taskpane; taskpane loads React app at `https://localhost:3000/taskpane.html`.

## Build
```bash
npm run build   # production bundle (replaces localhost URLs with urlProd in webpack.config.js)
npm run validate
```

Notes:
- React app lives in `src/taskpane/` (App.tsx, api.ts, office.ts, types.ts).
- Office commands untouched from the template. Update icons/resources in `assets/` if needed.
