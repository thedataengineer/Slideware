import { DeckSnapshot, SnapshotShape } from "./snapshot";

export interface Brand {
  headingFont: string;
  bodyFont: string;
  colors: string[];
}

export interface BrandFormat {
  id: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
}

const TEXT_SHAPE_TYPES = new Set([
  "GeometricShape",
  "TextBox",
  "Placeholder",
  "Freeform",
  "Callout",
]);

export function defaultBrand(): Brand {
  return {
    headingFont: "Segoe UI",
    bodyFont: "Segoe UI",
    colors: ["#20262e", "#0f6cbd", "#0b6a3c", "#b4009e", "#ffb900", "#f5f7f8"],
  };
}

export function serializeBrand(brand: Brand): string {
  return JSON.stringify(brand);
}

export function parseBrand(raw: string | null): Brand {
  if (!raw) return defaultBrand();
  try {
    const parsed = JSON.parse(raw) as Partial<Brand>;
    if (
      typeof parsed.headingFont !== "string" ||
      typeof parsed.bodyFont !== "string" ||
      !Array.isArray(parsed.colors) ||
      parsed.colors.length !== 6 ||
      parsed.colors.some((color) => typeof color !== "string")
    ) {
      return defaultBrand();
    }
    return { headingFont: parsed.headingFont, bodyFont: parsed.bodyFont, colors: parsed.colors };
  } catch {
    return defaultBrand();
  }
}

export function normalizeHex(value: string): string {
  const raw = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(raw)) return raw;
  if (/^#[0-9a-f]{3}$/.test(raw)) {
    return `#${raw[1]}${raw[1]}${raw[2]}${raw[2]}${raw[3]}${raw[3]}`;
  }
  throw new Error(`"${value}" is not a valid hex color.`);
}

function hasText(shape: SnapshotShape): boolean {
  return TEXT_SHAPE_TYPES.has(shape.type) && shape.text.trim().length > 0;
}

export function brandSelectionFormats(shapes: SnapshotShape[], brand: Brand): BrandFormat[] {
  return shapes
    .filter(hasText)
    .map((shape) => ({ id: shape.id, fontName: brand.bodyFont, fontColor: brand.colors[0] }));
}

export function brandDeckFormats(deck: DeckSnapshot, brand: Brand): BrandFormat[] {
  const formats: BrandFormat[] = [];
  deck.slides.forEach((slide) => {
    slide.shapes.filter(hasText).forEach((shape) => {
      formats.push({ id: shape.id, fontName: brand.bodyFont });
    });
  });
  return formats;
}
