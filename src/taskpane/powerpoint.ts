import { ShapeBounds } from "./alignment";
import { ShapeUpdate } from "./features/smartbar";
import { DeckSnapshot, SnapshotShape, SnapshotSlide, deriveTitle } from "./features/snapshot";

/* global Office, PowerPoint */

export type LayoutComputer = (shapes: ShapeBounds[]) => ShapeUpdate[];

export interface ShapeFormat {
  id: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
}

export interface InsertShapeSpec {
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

export class SmartAlignmentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SmartAlignmentError";
  }
}

class LayoutComputationFailure {
  constructor(readonly cause: unknown) {}
}

const TEXT_SHAPE_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Freeform", "Callout"]);

function assertSupported(): void {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    throw new SmartAlignmentError("Slideware requires PowerPoint API 1.5.");
  }
}

function wrapError(error: unknown, fallback: string): Error {
  if (error instanceof SmartAlignmentError) return error;
  if (error instanceof LayoutComputationFailure && error.cause instanceof Error) return error.cause;
  if (error instanceof Error && error.name !== "RichApi.Error") return error;
  return new SmartAlignmentError(fallback);
}

export async function applyLayout(compute: LayoutComputer): Promise<number> {
  assertSupported();

  try {
    return await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedShapes();
      selected.load("items");
      await context.sync();

      selected.items.forEach((shape) => shape.load("id,left,top,width,height"));
      await context.sync();

      const bounds: ShapeBounds[] = selected.items.map((shape) => ({
        id: shape.id,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }));
      let updates: ShapeUpdate[];
      try {
        updates = compute(bounds);
      } catch (error) {
        throw new LayoutComputationFailure(error);
      }
      const updateById = new Map(updates.map((update) => [update.id, update]));

      selected.items.forEach((shape) => {
        const update = updateById.get(shape.id);
        if (!update) return;
        if (update.left !== undefined) shape.left = update.left;
        if (update.top !== undefined) shape.top = update.top;
        if (update.width !== undefined) shape.width = update.width;
        if (update.height !== undefined) shape.height = update.height;
      });

      await context.sync();
      return updates.length;
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not update selected shapes.");
  }
}

interface LoadedShape {
  shape: PowerPoint.Shape;
  hasTextCapability: boolean;
}

function queueShapeLoads(shapes: PowerPoint.Shape[]): void {
  shapes.forEach((shape) => shape.load("id,name,type,left,top,width,height"));
}

function queueDetailLoads(shapes: PowerPoint.Shape[]): LoadedShape[] {
  return shapes.map((shape) => {
    const hasTextCapability = TEXT_SHAPE_TYPES.has(String(shape.type));
    if (hasTextCapability) {
      shape.textFrame.textRange.load("text");
      shape.textFrame.textRange.font.load("name,size,color");
      shape.fill.load("foregroundColor");
    }
    return { shape, hasTextCapability };
  });
}

function toSnapshotShape(loaded: LoadedShape): SnapshotShape {
  const { shape, hasTextCapability } = loaded;
  let text = "";
  let fillColor: string | undefined;
  let fontName: string | undefined;
  let fontSize: number | undefined;
  if (hasTextCapability) {
    try {
      text = shape.textFrame.textRange.text ?? "";
      fillColor = shape.fill.foregroundColor ?? undefined;
      fontName = shape.textFrame.textRange.font.name ?? undefined;
      fontSize = shape.textFrame.textRange.font.size ?? undefined;
    } catch {
      text = "";
    }
  }
  return {
    id: shape.id,
    name: shape.name,
    type: String(shape.type),
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
    text,
    fillColor,
    fontName,
    fontSize,
  };
}

export async function snapshotDeck(): Promise<DeckSnapshot> {
  assertSupported();

  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      slides.items.forEach((slide) => {
        slide.load("id");
        slide.shapes.load("items");
      });
      await context.sync();

      slides.items.forEach((slide) => queueShapeLoads(slide.shapes.items));
      await context.sync();

      const loadedBySlide = slides.items.map((slide) => queueDetailLoads(slide.shapes.items));
      await context.sync();

      const snapshotSlides: SnapshotSlide[] = slides.items.map((slide, index) => {
        const shapes = loadedBySlide[index].map(toSnapshotShape);
        return { id: slide.id, index: index + 1, title: deriveTitle(shapes), shapes };
      });

      return { slideCount: snapshotSlides.length, slides: snapshotSlides };
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not read the presentation.");
  }
}

export async function readSelection(): Promise<SnapshotShape[]> {
  assertSupported();

  try {
    return await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedShapes();
      selected.load("items");
      await context.sync();

      queueShapeLoads(selected.items);
      await context.sync();

      const loaded = queueDetailLoads(selected.items);
      await context.sync();

      return loaded.map(toSnapshotShape);
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not read the selection.");
  }
}

export async function writeShapeFormats(formats: ShapeFormat[]): Promise<number> {
  assertSupported();
  if (formats.length === 0) return 0;

  try {
    return await PowerPoint.run(async (context) => {
      const formatById = new Map(formats.map((format) => [format.id, format]));
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      slides.items.forEach((slide) => slide.shapes.load("items"));
      await context.sync();

      const allShapes: PowerPoint.Shape[] = [];
      slides.items.forEach((slide) => allShapes.push(...slide.shapes.items));
      allShapes.forEach((shape) => shape.load("id,type"));
      await context.sync();

      let applied = 0;
      allShapes.forEach((shape) => {
        const format = formatById.get(shape.id);
        if (!format) return;
        applied += 1;
        if (format.fillColor) shape.fill.setSolidColor(format.fillColor);
        if (TEXT_SHAPE_TYPES.has(String(shape.type))) {
          const font = shape.textFrame.textRange.font;
          if (format.fontName) font.name = format.fontName;
          if (format.fontSize) font.size = format.fontSize;
          if (format.fontColor) font.color = format.fontColor;
        }
      });
      await context.sync();
      return applied;
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not apply formatting.");
  }
}

async function currentSlide(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide> {
  const selectedSlides = context.presentation.getSelectedSlides();
  selectedSlides.load("items");
  await context.sync();
  if (selectedSlides.items.length > 0) return selectedSlides.items[0];

  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  if (slides.items.length === 0) {
    throw new SmartAlignmentError("The presentation has no slides.");
  }
  return slides.items[0];
}

export async function insertShapes(specs: InsertShapeSpec[]): Promise<number> {
  assertSupported();

  try {
    return await PowerPoint.run(async (context) => {
      const slide = await currentSlide(context);

      specs.forEach((spec) => {
        const options = { left: spec.left, top: spec.top, width: spec.width, height: spec.height };
        const shape =
          spec.kind === "textbox"
            ? slide.shapes.addTextBox(spec.text ?? "", options)
            : slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, options);
        if (spec.kind === "rect" && spec.text) {
          shape.textFrame.textRange.text = spec.text;
        }
        if (spec.fillColor) shape.fill.setSolidColor(spec.fillColor);
        if (spec.text || spec.kind === "textbox") {
          const font = shape.textFrame.textRange.font;
          if (spec.fontName) font.name = spec.fontName;
          if (spec.fontSize) font.size = spec.fontSize;
          if (spec.fontColor) font.color = spec.fontColor;
          if (spec.bold !== undefined) font.bold = spec.bold;
        }
      });

      await context.sync();
      return specs.length;
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not insert shapes.");
  }
}

export async function replaceShapeText(shapeId: string, text: string): Promise<void> {
  assertSupported();

  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      slides.items.forEach((slide) => slide.shapes.load("items"));
      await context.sync();

      const allShapes: PowerPoint.Shape[] = [];
      slides.items.forEach((slide) => allShapes.push(...slide.shapes.items));
      allShapes.forEach((shape) => shape.load("id"));
      await context.sync();

      const target = allShapes.find((shape) => shape.id === shapeId);
      if (!target) {
        throw new SmartAlignmentError("The shape is no longer in the presentation.");
      }
      target.textFrame.textRange.text = text;
      await context.sync();
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not replace the text.");
  }
}

export async function gotoSlide(slideIndex: number): Promise<void> {
  assertSupported();

  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      const slide = slides.items[slideIndex - 1];
      if (!slide) throw new SmartAlignmentError(`Slide ${slideIndex} does not exist.`);
      slide.load("id");
      await context.sync();

      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not switch slides.");
  }
}

export function canSelectShapes(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.6");
}

export async function setSelection(shapeIds: string[]): Promise<void> {
  assertSupported();
  if (!canSelectShapes()) {
    throw new SmartAlignmentError("Selecting shapes requires PowerPoint API 1.6.");
  }

  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation as unknown as {
        setSelectedShapes?: (ids: string[]) => void;
      };
      if (typeof presentation.setSelectedShapes !== "function") {
        throw new SmartAlignmentError("Selecting shapes requires PowerPoint API 1.6.");
      }
      presentation.setSelectedShapes(shapeIds);
      await context.sync();
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not select the shapes.");
  }
}
