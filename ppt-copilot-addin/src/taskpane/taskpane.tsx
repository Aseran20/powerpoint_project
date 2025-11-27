import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    return;
  }

  const rootElement = document.getElementById("root");
  if (!rootElement) {
    throw new Error("Root element not found");
  }

  createRoot(rootElement).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
});
