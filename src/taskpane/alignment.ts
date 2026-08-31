export interface ShapeBounds {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ShapePosition {
  id: string;
  left: number;
  top: number;
}

export type AlignMode = "left" | "center" | "right" | "top" | "middle" | "bottom";

export type DistributionAxis = "horizontal" | "vertical";

function requireShapes(shapes: ShapeBounds[], minimum: number): void {
  if (shapes.length < minimum) {
    throw new Error(`Select at least ${minimum} shapes.`);
  }
}

function selectionBounds(shapes: ShapeBounds[]) {
  return {
    left: Math.min(...shapes.map((shape) => shape.left)),
    top: Math.min(...shapes.map((shape) => shape.top)),
    right: Math.max(...shapes.map((shape) => shape.left + shape.width)),
    bottom: Math.max(...shapes.map((shape) => shape.top + shape.height)),
  };
}

export function alignShapes(shapes: ShapeBounds[], mode: AlignMode): ShapePosition[] {
  requireShapes(shapes, 2);
  const bounds = selectionBounds(shapes);
  const centerX = (bounds.left + bounds.right) / 2;
  const centerY = (bounds.top + bounds.bottom) / 2;

  return shapes.map((shape) => {
    let left = shape.left;
    let top = shape.top;

    if (mode === "left") left = bounds.left;
    if (mode === "center") left = centerX - shape.width / 2;
    if (mode === "right") left = bounds.right - shape.width;
    if (mode === "top") top = bounds.top;
    if (mode === "middle") top = centerY - shape.height / 2;
    if (mode === "bottom") top = bounds.bottom - shape.height;

    return { id: shape.id, left, top };
  });
}

export function distributeShapes(shapes: ShapeBounds[], axis: DistributionAxis): ShapePosition[] {
  requireShapes(shapes, 3);
  const horizontal = axis === "horizontal";
  const sorted = shapes
    .map((shape, index) => ({ shape, index }))
    .sort((a, b) => {
      const aCenter = horizontal
        ? a.shape.left + a.shape.width / 2
        : a.shape.top + a.shape.height / 2;
      const bCenter = horizontal
        ? b.shape.left + b.shape.width / 2
        : b.shape.top + b.shape.height / 2;
      return aCenter - bCenter || a.index - b.index;
    });
  const first = sorted[0].shape;
  const last = sorted[sorted.length - 1].shape;
  const start = horizontal ? first.left : first.top;
  const end = horizontal ? last.left + last.width : last.top + last.height;
  const totalSize = sorted.reduce(
    (sum, item) => sum + (horizontal ? item.shape.width : item.shape.height),
    0
  );
  const gap = (end - start - totalSize) / (sorted.length - 1);
  const positions = new Map<string, ShapePosition>();
  let cursor = start;

  sorted.forEach(({ shape }) => {
    positions.set(shape.id, {
      id: shape.id,
      left: horizontal ? cursor : shape.left,
      top: horizontal ? shape.top : cursor,
    });
    cursor += (horizontal ? shape.width : shape.height) + gap;
  });

  return shapes.map((shape) => positions.get(shape.id) as ShapePosition);
}

export interface MatrixOptions {
  columns: number;
  horizontalGap: number;
  verticalGap: number;
}

function readingOrder(shapes: ShapeBounds[]): ShapeBounds[] {
  return shapes
    .map((shape, index) => ({ shape, index }))
    .sort((a, b) => a.shape.top - b.shape.top || a.shape.left - b.shape.left || a.index - b.index)
    .map(({ shape }) => shape);
}

export function arrangeMatrix(shapes: ShapeBounds[], options: MatrixOptions): ShapePosition[] {
  requireShapes(shapes, 2);
  if (!Number.isInteger(options.columns) || options.columns < 1) {
    throw new Error("Columns must be a positive integer.");
  }
  if (options.columns > shapes.length) {
    throw new Error("Columns cannot exceed selected shape count.");
  }
  if (!Number.isFinite(options.horizontalGap) || !Number.isFinite(options.verticalGap)) {
    throw new Error("Matrix gaps must be valid numbers.");
  }

  const ordered = readingOrder(shapes);
  const rows = Math.ceil(ordered.length / options.columns);
  const columnWidths = Array.from({ length: options.columns }, () => 0);
  const rowHeights = Array.from({ length: rows }, () => 0);
  ordered.forEach((shape, index) => {
    const column = index % options.columns;
    const row = Math.floor(index / options.columns);
    columnWidths[column] = Math.max(columnWidths[column], shape.width);
    rowHeights[row] = Math.max(rowHeights[row], shape.height);
  });

  const bounds = selectionBounds(shapes);
  const columnLefts = columnWidths.map(
    (_, column) =>
      bounds.left +
      columnWidths.slice(0, column).reduce((sum, width) => sum + width, 0) +
      column * options.horizontalGap
  );
  const rowTops = rowHeights.map(
    (_, row) =>
      bounds.top +
      rowHeights.slice(0, row).reduce((sum, height) => sum + height, 0) +
      row * options.verticalGap
  );
  const positions = new Map<string, ShapePosition>();

  ordered.forEach((shape, index) => {
    const column = index % options.columns;
    const row = Math.floor(index / options.columns);
    positions.set(shape.id, {
      id: shape.id,
      left: columnLefts[column] + (columnWidths[column] - shape.width) / 2,
      top: rowTops[row] + (rowHeights[row] - shape.height) / 2,
    });
  });

  return shapes.map((shape) => positions.get(shape.id) as ShapePosition);
}

export interface CircleOptions {
  radius: number;
  startAngle: number;
  clockwise: boolean;
}

function cleanCoordinate(value: number): number {
  return Math.abs(value) < 1e-10 ? 0 : Number(value.toFixed(10));
}

export function arrangeCircle(shapes: ShapeBounds[], options: CircleOptions): ShapePosition[] {
  requireShapes(shapes, 2);
  if (!Number.isFinite(options.radius) || options.radius <= 0) {
    throw new Error("Radius must be greater than 0.");
  }
  if (!Number.isFinite(options.startAngle)) {
    throw new Error("Start angle must be a valid number.");
  }

  const ordered = readingOrder(shapes);
  const bounds = selectionBounds(shapes);
  const centerX = (bounds.left + bounds.right) / 2;
  const centerY = (bounds.top + bounds.bottom) / 2;
  const direction = options.clockwise ? 1 : -1;
  const positions = new Map<string, ShapePosition>();

  ordered.forEach((shape, index) => {
    const degrees = options.startAngle + (direction * (360 * index)) / ordered.length;
    const radians = (degrees * Math.PI) / 180;
    positions.set(shape.id, {
      id: shape.id,
      left: cleanCoordinate(centerX + options.radius * Math.cos(radians) - shape.width / 2),
      top: cleanCoordinate(centerY + options.radius * Math.sin(radians) - shape.height / 2),
    });
  });

  return shapes.map((shape) => positions.get(shape.id) as ShapePosition);
}
