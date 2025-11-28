/// <reference types="office-js" />

import { markdownToFlat } from "./markdownParser";

function getPowerPoint() {
    const ppt = (window as any).PowerPoint;
    if (!ppt || typeof ppt.run !== "function") {
        throw new Error(
            "L'API PowerPoint n'est pas prête. Ouvre le taskpane dans PowerPoint et attends quelques secondes."
        );
    }
    return ppt;
}

/**
 * Lire le texte de la première shape sélectionnée
 */
export async function getSelectedShapeText(): Promise<{ shapeId: string | null; text: string }> {
    const PowerPoint = getPowerPoint();

    return PowerPoint.run(async (context: any) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            return { shapeId: null, text: "" };
        }

        const shape = selectedShapes.items[0];
        shape.load("id, textFrame/textRange/text");
        await context.sync();

        const shapeId: string = shape.id;
        const text: string = shape.textFrame.textRange.text || "";

        return { shapeId, text };
    });
}

/**
 * Remplace le texte de la shape sélectionnée par du texte formaté
 * à partir de Markdown (markdownToFlat -> plain text + spans + bullets)
 */
export async function setFormattedText(markdownText: string): Promise<{ shapeId: string | null }> {
    const PowerPoint = getPowerPoint();
    const flattened = markdownToFlat(markdownText);

    return PowerPoint.run(async (context: any) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            return { shapeId: null };
        }

        const shape = selectedShapes.items[0];
        shape.load("id, textFrame/textRange");
        await context.sync();

        const textRange = shape.textFrame.textRange;

        // 1) Injecter le plain text généré par le parser
        textRange.text = flattened.text;
        await context.sync();

        // Recharger le texte pour connaître la longueur réelle
        textRange.load("text");
        await context.sync();

        const fullText: string = textRange.text || "";
        const totalLength = fullText.length;

        // 2) Appliquer les styles (gras, italique, headings) via getSubstring
        for (const span of flattened.spans) {
            try {
                if (span.length <= 0 || span.start < 0 || span.start >= totalLength) {
                    console.warn(
                        `Skipping style span out of bounds: start=${span.start}, length=${span.length}, total=${totalLength}`
                    );
                    continue;
                }

                const maxLength = totalLength - span.start;
                const safeLength = Math.min(span.length, maxLength);
                if (safeLength <= 0) {
                    console.warn(
                        `Skipping style span with non-positive safeLength: start=${span.start}, length=${span.length}, total=${totalLength}`
                    );
                    continue;
                }

                const subRange = textRange.getSubstring(span.start, safeLength);

                if (span.bold !== undefined) {
                    subRange.font.bold = span.bold;
                }
                if (span.italic !== undefined) {
                    subRange.font.italic = span.italic;
                }

                if (span.headingLevel) {
                    switch (span.headingLevel) {
                        case 1:
                            subRange.font.size = 28;
                            break;
                        case 2:
                            subRange.font.size = 24;
                            break;
                        case 3:
                            subRange.font.size = 20;
                            break;
                    }
                }

                await context.sync();
            } catch (err) {
                console.warn(`Could not apply style span at ${span.start}:`, err);
            }
        }

        // 3) Appliquer les bullets sur les blocs de listes
        for (const block of flattened.bullets) {
            try {
                if (block.length <= 0 || block.start < 0 || block.start >= totalLength) {
                    console.warn(
                        `Skipping bullet block out of bounds: start=${block.start}, length=${block.length}, total=${totalLength}`
                    );
                    continue;
                }

                const maxLength = totalLength - block.start;
                const safeLength = Math.min(block.length, maxLength);
                if (safeLength <= 0) {
                    console.warn(
                        `Skipping bullet block with non-positive safeLength: start=${block.start}, length=${block.length}, total=${totalLength}`
                    );
                    continue;
                }

                const listRange = textRange.getSubstring(block.start, safeLength);
                listRange.paragraphFormat.bulletFormat.visible = true;

                await context.sync();
            } catch (err) {
                console.warn(`Could not apply bullet at ${block.start}:`, err);
            }
        }

        return { shapeId: shape.id };
    });
}

/**
 * Compatibilité descendante
 * @deprecated Utiliser setFormattedText
 */
export async function setSelectedShapeText(newText: string): Promise<{ shapeId: string | null }> {
    return setFormattedText(newText);
}
