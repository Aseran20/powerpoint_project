import fs from "fs";
import os from "os";
import path from "path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const certPath = path.join(certDir, "localhost.crt");
const keyPath = path.join(certDir, "localhost.key");

export default defineConfig({
  plugins: [basicSsl(), react()],
  server: {
    port: 5173,
    host: "0.0.0.0",
    strictPort: true,
    https: fs.existsSync(certPath) && fs.existsSync(keyPath)
      ? {
          cert: fs.readFileSync(certPath),
          key: fs.readFileSync(keyPath),
        }
      : true,
  },
});
