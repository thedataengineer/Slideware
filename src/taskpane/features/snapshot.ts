export interface SnapshotShape {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text: string;
  fillColor?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
}

export interface SnapshotSlide {
  id: string;
  index: number;
  title?: string;
  shapes: SnapshotShape[];
}

export interface DeckSnapshot {
  slideCount: number;
  slides: SnapshotSlide[];
}

export interface TitleCandidate {
  name: string;
  top: number;
  text: string;
}

const TITLE_NAME = /(^|[^a-z])title([^a-z]|$)/i;
const SUBTITLE_NAME = /sub-?title/i;

export function deriveTitle(shapes: TitleCandidate[]): string | undefined {
  const withText = shapes.filter((shape) => shape.text.trim().length > 0);
  if (withText.length === 0) return undefined;

  // "Subtitle" must not satisfy this, and the topmost title wins over document order.
  const titled = withText
    .filter((shape) => TITLE_NAME.test(shape.name) && !SUBTITLE_NAME.test(shape.name))
    .sort((a, b) => a.top - b.top);
  if (titled.length > 0) return titled[0].text.trim();

  const topmost = withText.reduce((best, shape) => (shape.top < best.top ? shape : best));
  return topmost.text.trim();
}
