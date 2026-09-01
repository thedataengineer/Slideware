import { DeckSnapshot } from "./snapshot";

export interface FontUsage {
  name: string;
  count: number;
}

export interface FontFormat {
  id: string;
  fontName: string;
}

export function usedFonts(deck: DeckSnapshot): FontUsage[] {
  const counts = new Map<string, number>();
  deck.slides.forEach((slide) => {
    slide.shapes.forEach((shape) => {
      if (shape.text.trim().length === 0 || !shape.fontName) return;
      counts.set(shape.fontName, (counts.get(shape.fontName) ?? 0) + 1);
    });
  });
  return Array.from(counts.entries())
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
}

export function fontReplaceFormats(deck: DeckSnapshot, from: string, to: string): FontFormat[] {
  const source = from.trim().toLowerCase();
  const target = to.trim();
  if (source.length === 0) throw new Error("Pick which font to replace.");
  if (target.length === 0) throw new Error("Pick the replacement font.");

  const formats: FontFormat[] = [];
  deck.slides.forEach((slide) => {
    slide.shapes.forEach((shape) => {
      if (shape.fontName?.toLowerCase() === source) {
        formats.push({ id: shape.id, fontName: target });
      }
    });
  });

  if (formats.length === 0) {
    throw new Error(`No shapes use "${from.trim()}".`);
  }
  return formats;
}
