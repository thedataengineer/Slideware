export interface ShapeDescriptor {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  /** Measured by probing textFrame, not inferred from the shape type. */
  canHoldText: boolean;
  /** Only populated when the host is PowerPointApi 1.8 or newer. */
  placeholderType?: string;
}

export interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface PlaceholderPick {
  titleShapeId?: string;
  bodyShapeId?: string;
}

export interface SlideSize {
  width: number;
  height: number;
}

const MARGIN = 60;
const TITLE_TYPES = ["Title", "CenterTitle"];
const BODY_TYPES = ["Body", "Object", "Content"];
// "Subtitle" must not satisfy the title rule, which is the bug deriveTitle used to have.
const TITLE_NAME = /(^|[^a-z])title([^a-z]|$)/i;
const SUBTITLE_NAME = /sub-?title/i;
const BODY_NAME = /content|body|text/i;
const NON_TEXT_NAME = /picture|image|chart|table|diagram|smart\s*art|media|icon/i;

function byPosition(a: ShapeDescriptor, b: ShapeDescriptor): number {
  return a.top - b.top || a.left - b.left || a.id.localeCompare(b.id);
}

function byArea(a: ShapeDescriptor, b: ShapeDescriptor): number {
  return b.width * b.height - a.width * a.height || byPosition(a, b);
}

export function pickPlaceholders(shapes: ShapeDescriptor[]): PlaceholderPick {
  const usable = shapes.filter((shape) => shape.canHoldText);
  if (usable.length === 0) return {};

  const typedTitle = usable
    .filter((shape) => shape.placeholderType && TITLE_TYPES.includes(shape.placeholderType))
    .sort(byPosition)[0];
  const namedTitle = usable
    .filter((shape) => TITLE_NAME.test(shape.name) && !SUBTITLE_NAME.test(shape.name))
    .sort(byPosition)[0];
  const title = typedTitle ?? namedTitle ?? usable.slice().sort(byPosition)[0];

  const rest = usable.filter((shape) => shape.id !== title.id);
  const typedBody = rest
    .filter((shape) => shape.placeholderType && BODY_TYPES.includes(shape.placeholderType))
    .sort(byArea)[0];
  const namedBody = rest
    .filter((shape) => BODY_NAME.test(shape.name) && !NON_TEXT_NAME.test(shape.name))
    .sort(byArea)[0];
  const anyBody = rest
    .filter((shape) => !NON_TEXT_NAME.test(shape.name) && shape.top >= title.top)
    .sort(byArea)[0];

  const body = typedBody ?? namedBody ?? anyBody;
  return { titleShapeId: title.id, bodyShapeId: body?.id };
}

/**
 * Where to draw text when the layout gave no usable placeholder. The layout's own geometry is
 * preferred so a drawn box still lands where the template puts text.
 */
export function fallbackRects(
  layoutShapes: ShapeDescriptor[],
  slideSize: SlideSize
): { title: Rect; body: Rect } {
  const pick = pickPlaceholders(layoutShapes);
  const byId = new Map(layoutShapes.map((shape) => [shape.id, shape]));
  const titleShape = pick.titleShapeId ? byId.get(pick.titleShapeId) : undefined;
  const bodyShape = pick.bodyShapeId ? byId.get(pick.bodyShapeId) : undefined;

  const contentWidth = slideSize.width - MARGIN * 2;
  const defaultTitle: Rect = { left: MARGIN, top: MARGIN, width: contentWidth, height: 70 };
  const defaultBody: Rect = {
    left: MARGIN,
    top: MARGIN + 100,
    width: contentWidth,
    height: slideSize.height - MARGIN * 2 - 100,
  };

  return {
    title: titleShape ? toRect(titleShape) : defaultTitle,
    body: bodyShape ? toRect(bodyShape) : defaultBody,
  };
}

function toRect(shape: ShapeDescriptor): Rect {
  return { left: shape.left, top: shape.top, width: shape.width, height: shape.height };
}
