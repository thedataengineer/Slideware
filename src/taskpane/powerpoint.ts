import { ShapeBounds, ShapePosition } from "./alignment";

/* global Office, PowerPoint */

export type LayoutComputer = (shapes: ShapeBounds[]) => ShapePosition[];

export class SmartAlignmentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SmartAlignmentError";
  }
}

class LayoutComputationFailure {
  constructor(readonly cause: unknown) {}
}

function assertSupported(): void {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    throw new SmartAlignmentError("Smart Alignment requires PowerPoint API 1.5.");
  }
}

export async function applyLayout(compute: LayoutComputer): Promise<number> {
  assertSupported();

  try {
    return await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedShapes();
      selected.load("items");
      await context.sync();

      selected.items.forEach((shape) => shape.load("id,left,top,width,height"));
      await context.sync();

      const bounds: ShapeBounds[] = selected.items.map((shape) => ({
        id: shape.id,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }));
      let positions: ShapePosition[];
      try {
        positions = compute(bounds);
      } catch (error) {
        throw new LayoutComputationFailure(error);
      }
      const positionById = new Map(positions.map((position) => [position.id, position]));

      const updates = selected.items.map((shape) => {
        const position = positionById.get(shape.id);
        if (!position) {
          throw new SmartAlignmentError(`Layout returned no position for shape ${shape.id}.`);
        }
        return { shape, position };
      });

      updates.forEach(({ shape, position }) => {
        shape.left = position.left;
        shape.top = position.top;
      });

      await context.sync();
      return selected.items.length;
    });
  } catch (error) {
    if (error instanceof SmartAlignmentError) {
      throw error;
    }
    if (error instanceof LayoutComputationFailure && error.cause instanceof Error) {
      throw error.cause;
    }
    throw new SmartAlignmentError("PowerPoint could not update selected shapes.");
  }
}
