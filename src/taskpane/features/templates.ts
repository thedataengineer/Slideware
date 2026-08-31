import { Brand } from "./branding";

export type TemplateName = "title-block" | "kpi-row" | "quote-card" | "section-divider";

export interface SlideSizeSpec {
  width: number;
  height: number;
}

export interface TemplateShapeSpec {
  kind: "textbox" | "rect";
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  fillColor?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
}

export function templateShapes(
  name: string,
  brand: Brand,
  slideSize: SlideSizeSpec
): TemplateShapeSpec[] {
  const margin = 60;
  const contentWidth = slideSize.width - margin * 2;

  if (name === "title-block") {
    return [
      {
        kind: "textbox",
        left: margin,
        top: slideSize.height * 0.3,
        width: contentWidth,
        height: 60,
        text: "Presentation title",
        fontName: brand.headingFont,
        fontSize: 40,
        fontColor: brand.colors[0],
        bold: true,
      },
      {
        kind: "textbox",
        left: margin,
        top: slideSize.height * 0.3 + 70,
        width: contentWidth,
        height: 30,
        text: "Subtitle or presenter name",
        fontName: brand.bodyFont,
        fontSize: 20,
        fontColor: brand.colors[1],
      },
    ];
  }

  if (name === "kpi-row") {
    const gap = 20;
    const tileWidth = (contentWidth - gap * 2) / 3;
    const top = slideSize.height * 0.35;
    return [0, 1, 2].map((index) => ({
      kind: "rect" as const,
      left: margin + index * (tileWidth + gap),
      top,
      width: tileWidth,
      height: 120,
      text: "0%\nKPI label",
      fillColor: brand.colors[5],
      fontName: brand.bodyFont,
      fontSize: 18,
      fontColor: brand.colors[0],
    }));
  }

  if (name === "quote-card") {
    return [
      {
        kind: "rect",
        left: margin,
        top: slideSize.height * 0.25,
        width: contentWidth,
        height: 200,
        text: '"A memorable quote goes here."\n— Attribution',
        fillColor: brand.colors[5],
        fontName: brand.headingFont,
        fontSize: 24,
        fontColor: brand.colors[0],
      },
    ];
  }

  if (name === "section-divider") {
    return [
      {
        kind: "rect",
        left: 0,
        top: slideSize.height * 0.4,
        width: slideSize.width,
        height: 110,
        fillColor: brand.colors[1],
      },
      {
        kind: "textbox",
        left: margin,
        top: slideSize.height * 0.4 + 30,
        width: contentWidth,
        height: 50,
        text: "Section title",
        fontName: brand.headingFont,
        fontSize: 32,
        fontColor: "#ffffff",
        bold: true,
      },
    ];
  }

  throw new Error(`Unknown template: ${name}.`);
}
