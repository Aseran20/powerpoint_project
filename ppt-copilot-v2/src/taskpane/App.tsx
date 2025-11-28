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
import type { ChatMessage } from "./types";
import { MarkdownText } from "./MarkdownText";

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

    const handleSend = async () => {
        if (!input.trim()) return;
        setLoading(true);

        try {
            // 1. Get context from PowerPoint
            const { text: currentText } = await getSelectedShapeText();

            // 2. Add user message
            const userMsg: ChatMessage = {
                id: Date.now().toString(),
                role: "user",
                content: input,
                createdAt: Date.now()
            };

            const newMessages = [...messages, userMsg];
            setMessages(newMessages);
            setInput("");

            // 3. Prepare context for AI
            // Send only the current message with optional context (not full history)
            const messagesToSend: ChatMessage[] = currentText
                ? [
                    { id: "ctx", role: "user", content: `CONTEXTE SLIDE: ${currentText}\n\nINSTRUCTION UTILISATEUR: ${input}`, createdAt: 0 }
                ]
                : [userMsg];

            // 4. Call Backend
            const response = await sendChat(messagesToSend);

            // 5. Add assistant message
            const assistantMsg: ChatMessage = {
                id: (Date.now() + 1).toString(),
                role: "assistant",
                content: response.assistant_text,
                createdAt: Date.now()
            };
            setMessages([...newMessages, assistantMsg]);

            // 6. Update PowerPoint
            await setSelectedShapeText(response.assistant_text);

        } catch (error) {
            console.error(error);
            setMessages(prev => [...prev, {
                id: Date.now().toString(),
                role: "assistant",
                content: `Erreur: ${error}`,
                createdAt: Date.now()
            }]);
        } finally {
            setLoading(false);
        }
    };

    const handleUndo = async () => {
        // Implement undo logic if needed, for now just a placeholder
        console.log("Undo clicked");
    };

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
                                display: "inline-block",
                                textAlign: "left"
                            }}>
                                {msg.role === "assistant" ? (
                                    <MarkdownText text={msg.content} />
                                ) : (
                                    msg.content
                                )}
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
                            onClick={handleSend}
                            disabled={loading}
                            style={{ flexGrow: 1 }}
                        >
                            Générer
                        </Button>
                        <Button
                            appearance="subtle"
                            icon={<ArrowUndo24Regular />}
                            onClick={handleUndo}
                            title="Annuler la dernière action IA"
                        />
                    </div>
                </div>
            </div>
        </FluentProvider>
    );
};

export default App;
