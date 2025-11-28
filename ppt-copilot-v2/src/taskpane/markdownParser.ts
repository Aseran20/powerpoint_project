import { marked, Tokens } from "marked";

export type HeadingLevel = 1 | 2 | 3;

export interface StyleSpan {
    start: number;          // index dans plainText
    length: number;         // nombre de caractères
    bold?: boolean;
    italic?: boolean;
    headingLevel?: HeadingLevel;
}

export interface BulletBlock {
    start: number;          // index du début du bloc de liste (un item)
    length: number;         // nombre de caractères
}

export interface MarkdownFlattened {
    text: string;           // plainText final
    spans: StyleSpan[];
    bullets: BulletBlock[];
}

/**
 * Parse Markdown avec marked et produit:
 * - text: string "flat" (sans **, *...*)
 * - spans: segments à styler (gras, italique, heading)
 * - bullets: blocs à transformer en listes PowerPoint (un bloc = un item)
 */
export function markdownToFlat(md: string): MarkdownFlattened {
    const tokens = marked.lexer(md);

    let text = "";
    const spans: StyleSpan[] = [];
    const bullets: BulletBlock[] = [];

    let cursor = 0; // position actuelle dans `text`

    function appendText(str: string): number {
        text += str;
        const added = str.length;
        cursor += added;
        return added;
    }

    function addSpan(start: number, length: number, style: Partial<StyleSpan>) {
        if (length <= 0) return;
        spans.push({
            start,
            length,
            ...style,
        });
    }

    /**
     * Gestion des inline tokens (strong, em, text, link, codespan...)
     */
    function walkInline(inlineTokens?: Tokens.Generic[]): void {
        if (!inlineTokens) return;

        for (const token of inlineTokens) {
            switch (token.type) {
                case "strong": {
                    const t = token as Tokens.Strong;
                    const before = cursor;
                    walkInline(t.tokens || []);
                    const spanLength = cursor - before;
                    addSpan(before, spanLength, { bold: true });
                    break;
                }
                case "em": {
                    const t = token as Tokens.Em;
                    const before = cursor;
                    walkInline(t.tokens || []);
                    const spanLength = cursor - before;
                    addSpan(before, spanLength, { italic: true });
                    break;
                }
                case "codespan": {
                    const t = token as Tokens.Codespan;
                    appendText(t.text);
                    break;
                }
                case "link": {
                    const t = token as Tokens.Link;
                    const before = cursor;
                    walkInline(t.tokens || []);
                    // plus tard: style spécial pour les liens si besoin
                    break;
                }
                case "text": {
                    const t = token as Tokens.Text;
                    const nested = (t as any).tokens as Tokens.Generic[] | undefined;
                    if (nested && nested.length > 0) {
                        walkInline(nested);
                    } else {
                        appendText(t.text);
                    }
                    break;
                }
                case "space": {
                    appendText(" ");
                    break;
                }
                default: {
                    const anyToken = token as any;
                    if (typeof anyToken.text === "string") {
                        appendText(anyToken.text);
                    } else if (typeof anyToken.raw === "string") {
                        appendText(anyToken.raw);
                    }
                    break;
                }
            }
        }
    }

    /**
     * Parcours des tokens "block"
     */
    for (const token of tokens) {
        switch (token.type) {
            case "heading": {
                const h = token as Tokens.Heading;
                const start = cursor;
                walkInline(h.tokens || []);
                const length = cursor - start;

                const depth = Math.max(1, Math.min(3, h.depth)) as HeadingLevel;
                addSpan(start, length, { headingLevel: depth, bold: true });

                appendText("\n\n");
                break;
            }

            case "paragraph": {
                const p = token as Tokens.Paragraph;
                walkInline(p.tokens || []);
                appendText("\n\n");
                break;
            }

            case "list": {
                const list = token as Tokens.List;

                // On crée UN BulletBlock PAR item (et pas un seul pour toute la liste)
                for (const item of list.items) {
                    const li = item as Tokens.ListItem;

                    // Début de cet item dans le plain text
                    const itemStart = cursor;

                    if (li.tokens && li.tokens.length > 0) {
                        for (const child of li.tokens) {
                            if (child.type === "paragraph") {
                                const p = child as Tokens.Paragraph;
                                walkInline(p.tokens || []);
                            } else {
                                // fallback: traiter le child comme inline unique
                                walkInline([child as any]);
                            }
                        }
                    }

                    // Fin de l'item: on ajoute un retour ligne
                    appendText("\n");

                    const itemLength = cursor - itemStart; // inclut le \n
                    if (itemLength > 0) {
                        bullets.push({
                            start: itemStart,
                            length: itemLength,
                        });
                    }
                }

                // Ligne vide après la liste pour séparer du paragraphe suivant
                appendText("\n");
                break;
            }

            case "space": {
                appendText("\n");
                break;
            }

            default: {
                const anyToken = token as any;
                if (typeof anyToken.text === "string") {
                    appendText(anyToken.text);
                    appendText("\n\n");
                } else if (typeof anyToken.raw === "string") {
                    appendText(anyToken.raw);
                    appendText("\n\n");
                }
                break;
            }
        }
    }

    const result: MarkdownFlattened = { text, spans, bullets };

    // Debug
    console.log("=== Markdown Parser Debug ===");
    console.log("Raw markdown:", md);
    console.log("Text length:", text.length);
    console.log("Text preview:", text.substring(0, 200));
    console.log("Spans:", spans);
    console.log("Bullets:", bullets);
    console.log("============================");

    return result;
}

/* ------------------------------------------------------------------
 * Legacy exports pour le composant React MarkdownText
 * ------------------------------------------------------------------ */

export interface FormattedParagraph {
    segments: TextSegment[];
    isBullet: boolean;
}

export interface TextSegment {
    text: string;
    bold?: boolean;
    italic?: boolean;
}

/**
 * Legacy simple parser pour la UI (pas utilisé pour PowerPoint)
 */
export function parseMarkdown(markdown: string): FormattedParagraph[] {
    const lines = markdown.split("\n");
    const paragraphs: FormattedParagraph[] = [];

    for (const line of lines) {
        if (!line.trim()) continue;

        const bulletMatch = line.match(/^-\s+(.*)$/);
        const isBullet = bulletMatch !== null;
        const content = isBullet ? bulletMatch![1] : line;

        const segments = parseInlineFormatting(content);

        paragraphs.push({
            segments,
            isBullet,
        });
    }

    return paragraphs;
}

function parseInlineFormatting(text: string): TextSegment[] {
    const segments: TextSegment[] = [];
    let currentPos = 0;

    const regex = /(\*\*([^*]+)\*\*)|(\*([^*]+)\*)/g;
    let match: RegExpExecArray | null;

    while ((match = regex.exec(text)) !== null) {
        if (match.index > currentPos) {
            segments.push({
                text: text.substring(currentPos, match.index),
            });
        }

        if (match[1]) {
            segments.push({
                text: match[2],
                bold: true,
            });
        } else if (match[3]) {
            segments.push({
                text: match[4],
                italic: true,
            });
        }

        currentPos = regex.lastIndex;
    }

    if (currentPos < text.length) {
        segments.push({
            text: text.substring(currentPos),
        });
    }

    return segments;
}
