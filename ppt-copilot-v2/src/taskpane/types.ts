export type Role = "user" | "assistant";

export interface ChatMessage {
    id: string;
    role: Role;
    content: string;
    createdAt: number;
}

export interface UndoEntry {
    shapeId: string;
    previousText: string;
}

export interface UiState {
    messages: ChatMessage[];
    lastAssistantMessageId: string | null;
    lastUndoByShape: Record<string, UndoEntry | undefined>;
}
