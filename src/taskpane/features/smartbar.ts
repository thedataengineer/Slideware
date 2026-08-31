import { ShapeBounds } from "../alignment";

export type SizeMode = "width" | "height" | "both";

export interface ShapeUpdate {
  id: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export function matchSizes(shapes: ShapeBounds[], mode: SizeMode): ShapeUpdate[] {
  if (shapes.length < 2) {
    throw new Error("Select at least 2 shapes.");
  }
  const reference = shapes[0];

  return shapes.slice(1).map((shape) => {
    const update: ShapeUpdate = { id: shape.id };
    if (mode === "width" || mode === "both") update.width = reference.width;
    if (mode === "height" || mode === "both") update.height = reference.height;
    return update;
  });
}

export function swapPositions(shapes: ShapeBounds[]): ShapeUpdate[] {
  if (shapes.length !== 2) {
    throw new Error("Select exactly 2 shapes.");
  }
  const [first, second] = shapes;

  return [
    { id: first.id, left: second.left, top: second.top },
    { id: second.id, left: first.left, top: first.top },
  ];
}
