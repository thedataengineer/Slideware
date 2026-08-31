import { DeckSnapshot } from "./snapshot";

export interface SearchHit {
  slideIndex: number;
  shapeId: string;
  shapeName: string;
  snippet: string;
}

const SNIPPET_CONTEXT = 30;

export function searchDeck(deck: DeckSnapshot, query: string): SearchHit[] {
  const needle = query.trim().toLowerCase();
  if (needle.length === 0) {
    throw new Error("Type something to search for.");
  }

  const hits: SearchHit[] = [];
  deck.slides.forEach((slide) => {
    slide.shapes.forEach((shape) => {
      const haystack = shape.text.toLowerCase();
      const position = haystack.indexOf(needle);
      if (position === -1) return;

      const start = Math.max(0, position - SNIPPET_CONTEXT);
      const end = Math.min(shape.text.length, position + needle.length + SNIPPET_CONTEXT);
      const prefix = start > 0 ? "…" : "";
      const suffix = end < shape.text.length ? "…" : "";
      hits.push({
        slideIndex: slide.index,
        shapeId: shape.id,
        shapeName: shape.name,
        snippet: `${prefix}${shape.text.slice(start, end)}${suffix}`,
      });
    });
  });

  return hits;
}
