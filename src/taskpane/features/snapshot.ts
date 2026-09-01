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

export function deriveTitle(shapes: TitleCandidate[]): string | undefined {
  const withText = shapes.filter((shape) => shape.text.trim().length > 0);
  if (withText.length === 0) return undefined;

  const named = withText.find((shape) => shape.name.toLowerCase().includes("title"));
  if (named) return named.text.trim();

  const topmost = withText.reduce((best, shape) => (shape.top < best.top ? shape : best));
  return topmost.text.trim();
}
