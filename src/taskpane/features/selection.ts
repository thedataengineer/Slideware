import { SnapshotShape } from "./snapshot";

export interface SelectionCriteria {
  sameType?: boolean;
  sameFill?: boolean;
  sameSize?: boolean;
  sameFont?: boolean;
  sameFontColor?: boolean;
  sameFontSize?: boolean;
}

const SIZE_TOLERANCE = 1;

function equalIgnoringCase(a: string | undefined, b: string | undefined): boolean {
  return a !== undefined && b !== undefined && a.toLowerCase() === b.toLowerCase();
}

export function matchShapes(
  all: SnapshotShape[],
  anchor: SnapshotShape,
  criteria: SelectionCriteria
): string[] {
  if (!Object.values(criteria).some((flag) => flag === true)) {
    throw new Error("Pick at least one criteria.");
  }

  return all
    .filter((shape) => {
      if (shape.id === anchor.id) return true;
      if (criteria.sameType && shape.type !== anchor.type) return false;
      if (criteria.sameFill && !equalIgnoringCase(shape.fillColor, anchor.fillColor)) return false;
      if (criteria.sameSize) {
        if (
          Math.abs(shape.width - anchor.width) > SIZE_TOLERANCE ||
          Math.abs(shape.height - anchor.height) > SIZE_TOLERANCE
        ) {
          return false;
        }
      }
      if (criteria.sameFont && !equalIgnoringCase(shape.fontName, anchor.fontName)) return false;
      if (criteria.sameFontColor && !equalIgnoringCase(shape.fontColor, anchor.fontColor)) {
        return false;
      }
      if (criteria.sameFontSize) {
        if (shape.fontSize === undefined || anchor.fontSize === undefined) return false;
        if (Math.abs(shape.fontSize - anchor.fontSize) > 0.5) return false;
      }
      return true;
    })
    .map((shape) => shape.id);
}
