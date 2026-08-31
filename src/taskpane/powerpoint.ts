import { ShapeBounds } from "./alignment";
import { ShapeUpdate } from "./features/smartbar";

/* global Office, PowerPoint */

export type LayoutComputer = (shapes: ShapeBounds[]) => ShapeUpdate[];

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
      let updates: ShapeUpdate[];
      try {
        updates = compute(bounds);
      } catch (error) {
        throw new LayoutComputationFailure(error);
      }
      const updateById = new Map(updates.map((update) => [update.id, update]));

      selected.items.forEach((shape) => {
        const update = updateById.get(shape.id);
        if (!update) return;
        if (update.left !== undefined) shape.left = update.left;
        if (update.top !== undefined) shape.top = update.top;
        if (update.width !== undefined) shape.width = update.width;
        if (update.height !== undefined) shape.height = update.height;
      });

      await context.sync();
      return updates.length;
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
