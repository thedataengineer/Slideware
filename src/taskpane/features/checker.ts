import { DeckSnapshot } from "./snapshot";

export interface SlideSize {
  width: number;
  height: number;
}

export interface Finding {
  rule: "off-slide" | "tiny-font" | "font-sprawl" | "empty-text" | "overlong";
  message: string;
  slideIndex?: number;
  shapeId?: string;
  shapeName?: string;
}

const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
const MIN_FONT_SIZE = 12;
const MAX_FONT_FAMILIES = 3;
const MAX_TEXT_LENGTH = 300;
const TEXT_SHAPE_TYPES = new Set([
  "GeometricShape",
  "TextBox",
  "Placeholder",
  "Freeform",
  "Callout",
]);

export function audit(deck: DeckSnapshot, slideSize: SlideSize = DEFAULT_SLIDE_SIZE): Finding[] {
  const findings: Finding[] = [];
  const fontFamilies = new Set<string>();

  deck.slides.forEach((slide) => {
    slide.shapes.forEach((shape) => {
      const common = { slideIndex: slide.index, shapeId: shape.id, shapeName: shape.name };
      const hasText = shape.text.trim().length > 0;

      if (
        shape.left < 0 ||
        shape.top < 0 ||
        shape.left + shape.width > slideSize.width ||
        shape.top + shape.height > slideSize.height
      ) {
        findings.push({
          rule: "off-slide",
          message: `"${shape.name}" extends beyond the slide.`,
          ...common,
        });
      }

      if (hasText && shape.fontName) {
        fontFamilies.add(shape.fontName);
      }

      if (hasText && shape.fontSize !== undefined && shape.fontSize < MIN_FONT_SIZE) {
        findings.push({
          rule: "tiny-font",
          message: `"${shape.name}" uses ${shape.fontSize}pt text; keep body text at ${MIN_FONT_SIZE}pt or larger.`,
          ...common,
        });
      }

      if (!hasText && TEXT_SHAPE_TYPES.has(shape.type)) {
        findings.push({
          rule: "empty-text",
          message: `"${shape.name}" is an empty text shape.`,
          ...common,
        });
      }

      if (shape.text.length > MAX_TEXT_LENGTH) {
        findings.push({
          rule: "overlong",
          message: `"${shape.name}" holds ${shape.text.length} characters; split or tighten it.`,
          ...common,
        });
      }
    });
  });

  if (fontFamilies.size > MAX_FONT_FAMILIES) {
    findings.push({
      rule: "font-sprawl",
      message: `The deck uses ${fontFamilies.size} font families (${Array.from(fontFamilies).join(", ")}); trim to ${MAX_FONT_FAMILIES}.`,
    });
  }

  return findings;
}
