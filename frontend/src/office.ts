/// <reference types="office-js" />

function getPowerPoint() {
  const ppt = (window as any).PowerPoint;
  if (!ppt || typeof ppt.run !== "function") {
    throw new Error("L'API PowerPoint n'est pas prête. Ouvre le taskpane dans PowerPoint et attends quelques secondes.");
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
