import { SnapshotShape } from "./snapshot";

export interface SelectionCriteria {
  sameType?: boolean;
  sameFill?: boolean;
  sameSize?: boolean;
}

const SIZE_TOLERANCE = 1;

export function matchShapes(
  all: SnapshotShape[],
  anchor: SnapshotShape,
  criteria: SelectionCriteria
): string[] {
  if (!criteria.sameType && !criteria.sameFill && !criteria.sameSize) {
    throw new Error("Pick at least one criteria.");
  }

  return all
    .filter((shape) => {
      if (shape.id === anchor.id) return true;
      if (criteria.sameType && shape.type !== anchor.type) return false;
      if (criteria.sameFill) {
        const shapeFill = shape.fillColor?.toLowerCase();
        const anchorFill = anchor.fillColor?.toLowerCase();
        if (!shapeFill || !anchorFill || shapeFill !== anchorFill) return false;
      }
      if (criteria.sameSize) {
        if (
          Math.abs(shape.width - anchor.width) > SIZE_TOLERANCE ||
          Math.abs(shape.height - anchor.height) > SIZE_TOLERANCE
        ) {
          return false;
        }
      }
      return true;
    })
    .map((shape) => shape.id);
}
