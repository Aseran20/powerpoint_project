/// <reference types="office-js" />

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
 * Read the selected shape's text in PowerPoint.
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

    const shapeId = shape.id;
    const text = shape.textFrame.textRange.text || "";

    return { shapeId, text };
  });
}

/**
 * Replace the selected shape's text in PowerPoint.
 */
export async function setSelectedShapeText(newText: string): Promise<{ shapeId: string | null }> {
  const PowerPoint = getPowerPoint();
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

    shape.textFrame.textRange.text = newText;
    await context.sync();

    return { shapeId: shape.id };
  });
}

type BoldRange = { start: number; length: number };
type ParsedText = { plainText: string; boldRanges: BoldRange[]; bulletLines: boolean[] };

function parseMarkdownish(text: string): ParsedText {
  const boldRanges: BoldRange[] = [];
  const bulletLines: boolean[] = [];

  // Detect bullets: lines starting with "- "
  const lines = text.split(/\r?\n/);
  const processedLines: string[] = [];
  let cursor = 0;

  const boldPattern = /\*\*(.+?)\*\*/g;

  for (const line of lines) {
    const isBullet = line.trimStart().startsWith("- ");
    bulletLines.push(isBullet);

    // Remove leading "- " for bullets and prefix with a bullet char for display.
    const content = isBullet ? line.trimStart().slice(2) : line;
    let plainLine = "";
    let lastIndex = 0;

    // Strip **bold** markers while tracking positions.
    for (const match of content.matchAll(boldPattern)) {
      const [full, inner] = match;
      const startInLine = match.index ?? 0;
      plainLine += content.slice(lastIndex, startInLine);
      boldRanges.push({
        start: cursor + plainLine.length,
        length: inner.length,
      });
      plainLine += inner;
      lastIndex = startInLine + full.length;
    }

    plainLine += content.slice(lastIndex);

    // Prefix bullet character for readability (PowerPoint doesn't auto-bullet plain text).
    const finalLine = isBullet ? `• ${plainLine}` : plainLine;
    processedLines.push(finalLine);
    cursor += finalLine.length + 1; // +1 for the newline we will join with
  }

  return {
    plainText: processedLines.join("\n"),
    boldRanges,
    bulletLines,
  };
}

/**
 * Apply markdown-like formatting (bullets, **bold**) to the selected shape.
 */
export async function setSelectedShapeFormattedText(
  newText: string
): Promise<{ shapeId: string | null }> {
  const PowerPoint = getPowerPoint();
  const parsed = parseMarkdownish(newText);

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

    const range = shape.textFrame.textRange;
    range.text = parsed.plainText;
    await context.sync();

    // Apply bold ranges.
    for (const { start, length } of parsed.boldRanges) {
      try {
        const sub = range.getSubstring(start, length);
        sub.font.bold = true;
      } catch {
        // Ignore if substring not available.
      }
    }

    await context.sync();

    return { shapeId: shape.id };
  });
}
