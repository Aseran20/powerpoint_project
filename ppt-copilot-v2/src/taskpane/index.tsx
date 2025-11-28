import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./App";

/* global document, Office */

const title = "PPT Copilot";

const rootElement: HTMLElement | null = document.getElementById("container");

if (!rootElement) {
    throw new Error("Root element not found");
}

const root = createRoot(rootElement);

/* Render application after Office initializes */
Office.onReady(() => {
    root.render(
        <FluentProvider theme={webLightTheme}>
            <App />
        </FluentProvider>
    );
});
