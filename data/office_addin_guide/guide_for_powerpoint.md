📘 Plan de Refonte : PowerPoint Copilot (Standardisation)
Priorité : Stabiliser le projet en éliminant l'architecture hybride "Vite + Webpack" qui cause des bugs et des conflits de ports. Objectif : Migrer la logique existante dans un template Microsoft officiel propre, et adopter Fluent UI pour un look professionnel.

🛑 1. Diagnostic & Nettoyage
L'état actuel du projet contient deux structures en conflit :

❌ frontend/ (Vite) : Cause des problèmes de certificats et de build avec Office.

❌ ppt-copilot-addin/ (Webpack) : Configuration partielle.

Action : Nous allons ignorer ces dossiers et générer un nouveau dossier propre ppt-copilot-v2. Une fois la migration terminée, les anciens dossiers seront supprimés.

🛠️ 2. Génération du Socle Propre (CLI)
Générer le projet en utilisant le générateur officiel Yeoman avec les bons flags pour React.

Bash

# Se placer à la racine du repo
cd aseran20/powerpoint_project/

# Générer le nouveau projet standard
npx --package yo --package generator-office yo office \
  --projectType taskpane \
  --name ppt-copilot-v2 \
  --host powerpoint \
  --ts \
  --framework react \
  --skip-install
Note : Le flag --framework react est crucial pour avoir la structure App.tsx et le support JSX configuré nativement.

Ensuite, installer les dépendances et la librairie graphique Fluent UI v9 (le standard Office actuel) :

Bash

cd ppt-copilot-v2
npm install
npm install @fluentui/react-components
📦 3. Migration de la Logique Métier
Nous allons récupérer l'intelligence de l'ancien projet (fichiers api.ts, office.ts) et les mettre dans le nouveau.

A. Fichiers Helpers
Copier les fichiers suivants depuis l'ancien dossier frontend/src/ vers le nouveau ppt-copilot-v2/src/taskpane/ :

api.ts (Logique d'appel au backend Python)

office.ts (Manipulation de la slide PowerPoint)

types.ts (Interfaces TypeScript)

B. Configuration Backend
Vérifier dans src/taskpane/api.ts que l'URL du backend pointe bien vers votre serveur Python (ex: http://localhost:8000 ou l'URL du VPS).

🎨 4. Refonte de l'Interface (Fluent UI)
C'est l'étape clé pour ne plus avoir une UI "dégueu". On remplace le HTML brut par des composants Microsoft.

Fichier à modifier : ppt-copilot-v2/src/taskpane/App.tsx

Remplacer tout le contenu par ce modèle qui intègre votre logique existante avec le design system Office :

TypeScript

import * as React from "react";
import { 
  FluentProvider, 
  Button, 
  Textarea, 
  Body1, 
  Title3,
  webLightTheme, 
  makeStyles
} from "@fluentui/react-components";
import { Send24Regular, ArrowUndo24Regular } from "@fluentui/react-icons";
// Importer votre logique migrée
import { sendChat } from "./api";
import { getSelectedShapeText, setSelectedShapeText } from "./office";
import type { ChatMessage, UiState } from "./types";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "15px",
    padding: "15px",
    height: "100vh",
    boxSizing: "border-box",
  },
  chatWindow: {
    flexGrow: 1,
    border: "1px solid #e0e0e0",
    borderRadius: "8px",
    padding: "10px",
    overflowY: "auto",
    backgroundColor: "#fafafa",
  },
  inputArea: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  }
});

const App: React.FC = () => {
  const styles = useStyles();
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);

  // ... (Réintégrer ici la logique handleSend / handleApply de l'ancien App.tsx)

  return (
    <FluentProvider theme={webLightTheme}>
      <div className={styles.container}>
        <Title3>PPT Copilot</Title3>

        {/* Zone de Chat */}
        <div className={styles.chatWindow}>
          {messages.length === 0 && (
            <Body1 style={{ color: "#888", textAlign: "center", display: "block", marginTop: "20px" }}>
              Sélectionnez une forme et décrivez la modification souhaitée.
            </Body1>
          )}
          {messages.map((msg, i) => (
            <div key={i} style={{ marginBottom: "10px", textAlign: msg.role === "user" ? "right" : "left" }}>
              <span style={{ 
                background: msg.role === "user" ? "#0078d4" : "#e0e0e0", 
                color: msg.role === "user" ? "white" : "black",
                padding: "8px 12px", 
                borderRadius: "12px",
                display: "inline-block"
              }}>
                {msg.content}
              </span>
            </div>
          ))}
        </div>

        {/* Zone de Saisie */}
        <div className={styles.inputArea}>
          <Textarea 
            placeholder="Ex: Traduis ce texte en anglais..." 
            value={input}
            onChange={(e, data) => setInput(data.value)}
            resize="vertical"
          />
          
          <div style={{ display: "flex", gap: "10px" }}>
            <Button 
              appearance="primary" 
              icon={<Send24Regular />} 
              onClick={() => { /* Appel handleSend */ }}
              disabled={loading}
              style={{ flexGrow: 1 }}
            >
              Générer
            </Button>
            <Button 
              appearance="subtle"
              icon={<ArrowUndo24Regular />}
              onClick={() => { /* Appel handleUndo */ }}
              title="Annuler la dernière action IA"
            />
          </div>
        </div>
      </div>
    </FluentProvider>
  );
};

export default App;
🚀 5. Test et Validation
Lancer le backend Python (dans un terminal séparé) :

Bash

cd backend
source venv/bin/activate
python main.py
Lancer l'Add-in :

Bash

cd ppt-copilot-v2
npm run start
Cela va ouvrir PowerPoint Desktop automatiquement avec le panneau chargé.

Vérification :

Le panneau s'affiche-t-il avec le style Office (Fluent) ?

Les boutons sont-ils bleus ?

L'interaction avec la slide fonctionne-t-elle ?

🧹 6. Nettoyage Final
Une fois que ppt-copilot-v2 est validé :

Supprimer le dossier frontend (l'ancien code Vite).

Supprimer le dossier ppt-copilot-addin (l'ancien code Webpack mal configuré).

Renommer ppt-copilot-v2 en ppt-copilot-addin si désiré.