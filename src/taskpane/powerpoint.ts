import { ShapeBounds } from "./alignment";
import { loadInBatches } from "./features/batchload";
import {
  AddSlidesResult,
  PlannedSlide,
  SlideFailure,
  bodyTextFor,
  validatePlan,
} from "./features/deck-plan";
import { ShapeDescriptor, fallbackRects, pickPlaceholders } from "./features/placeholders";
import { DEFAULT_SLIDE_SIZE, SlideSize } from "./features/slide-size";
import { LayoutCatalog, LayoutRef, chooseLayouts } from "./features/slidelayout";
import { resolveSelectionTarget } from "./features/selection";
import { ShapeUpdate } from "./features/smartbar";
import { DeckSnapshot, SnapshotShape, SnapshotSlide, deriveTitle } from "./features/snapshot";

/* global Office, PowerPoint, console */

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

const TEXT_SHAPE_TYPES = new Set([
  "GeometricShape",
  "TextBox",
  "Placeholder",
  "Freeform",
  "Callout",
]);

function assertSupported(): void {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    throw new SmartAlignmentError("Slideware requires PowerPoint API 1.5.");
  }
}

interface HostErrorDebugInfo {
  code?: string;
  message?: string;
  errorLocation?: string;
  statement?: string;
}

function hostErrorDetail(error: unknown): string {
  const debugInfo = (error as { debugInfo?: HostErrorDebugInfo } | null)?.debugInfo;
  if (!debugInfo) return "";
  const code = debugInfo.code || (error as { code?: string }).code || "";
  const where = debugInfo.errorLocation || debugInfo.statement || "";
  return [code, where].filter(Boolean).join(" at ");
}

function wrapError(error: unknown, fallback: string): Error {
  if (error instanceof SmartAlignmentError) return error;
  if (error instanceof LayoutComputationFailure && error.cause instanceof Error) return error.cause;
  if (error instanceof Error && error.name !== "RichApi.Error") return error;
  const detail = hostErrorDetail(error);
  console.error(
    "Slideware host error",
    error,
    (error as { debugInfo?: HostErrorDebugInfo })?.debugInfo
  );
  return new SmartAlignmentError(detail ? `${fallback} (${detail})` : fallback);
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

function toCandidates(shapes: PowerPoint.Shape[]): LoadedShape[] {
  return shapes.map((shape) => ({
    shape,
    hasTextCapability: TEXT_SHAPE_TYPES.has(String(shape.type)),
  }));
}

function queueDetailLoad(candidate: LoadedShape): void {
  const { shape } = candidate;
  shape.textFrame.textRange.load("text");
  shape.textFrame.textRange.font.load("name,size,color");
  shape.fill.load("foregroundColor");
}

/**
 * Shape type alone does not prove a shape supports a text frame: PowerPoint throws
 * InvalidArgument at Shape.textFrame for picture placeholders and similar shapes, and that
 * rejects the entire sync. Isolate those shapes so one of them cannot blank a whole deck read.
 */
async function loadShapeDetails(
  context: PowerPoint.RequestContext,
  candidates: LoadedShape[]
): Promise<void> {
  const unsupported = await loadInBatches({
    items: candidates.filter((candidate) => candidate.hasTextCapability),
    queue: queueDetailLoad,
    sync: () => context.sync(),
  });
  unsupported.forEach((candidate) => {
    candidate.hasTextCapability = false;
  });
}

function toSnapshotShape(loaded: LoadedShape): SnapshotShape {
  const { shape, hasTextCapability } = loaded;
  let text = "";
  let fillColor: string | undefined;
  let fontName: string | undefined;
  let fontSize: number | undefined;
  let fontColor: string | undefined;
  if (hasTextCapability) {
    try {
      text = shape.textFrame.textRange.text ?? "";
      fillColor = shape.fill.foregroundColor ?? undefined;
      fontName = shape.textFrame.textRange.font.name ?? undefined;
      fontSize = shape.textFrame.textRange.font.size ?? undefined;
      fontColor = shape.textFrame.textRange.font.color ?? undefined;
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
    fontColor,
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

      const loadedBySlide = slides.items.map((slide) => toCandidates(slide.shapes.items));
      const allCandidates: LoadedShape[] = [];
      loadedBySlide.forEach((candidates) => allCandidates.push(...candidates));
      await loadShapeDetails(context, allCandidates);

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

      const loaded = toCandidates(selected.items);
      await loadShapeDetails(context, loaded);

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

export type TextFitMode = "text-to-box" | "box-to-text";

export async function setTextFit(shapeIds: string[], mode: TextFitMode): Promise<number> {
  assertSupported();
  if (shapeIds.length === 0) return 0;
  const setting = mode === "text-to-box" ? "TextToFitShape" : "ShapeToFitText";

  try {
    return await PowerPoint.run(async (context) => {
      const targets = new Set(shapeIds);
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
        if (!targets.has(shape.id) || !TEXT_SHAPE_TYPES.has(String(shape.type))) return;
        (shape.textFrame as unknown as { autoSizeSetting: string }).autoSizeSetting = setting;
        applied += 1;
      });
      await context.sync();
      return applied;
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not change the text autofit setting.");
  }
}

export async function deleteShapes(shapeIds: string[]): Promise<number> {
  assertSupported();
  if (shapeIds.length === 0) return 0;

  try {
    return await PowerPoint.run(async (context) => {
      const idsToDelete = new Set(shapeIds);
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      slides.items.forEach((slide) => slide.shapes.load("items"));
      await context.sync();

      const allShapes: PowerPoint.Shape[] = [];
      slides.items.forEach((slide) => allShapes.push(...slide.shapes.items));
      allShapes.forEach((shape) => shape.load("id"));
      await context.sync();

      let deleted = 0;
      allShapes.forEach((shape) => {
        if (idsToDelete.has(shape.id)) {
          shape.delete();
          deleted += 1;
        }
      });
      await context.sync();
      return deleted;
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not delete the original shapes.");
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
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
}

export async function setSelection(shapeIds: string[]): Promise<void> {
  assertSupported();
  if (shapeIds.length === 0) return;

  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      slides.items.forEach((slide) => {
        slide.load("id");
        slide.shapes.load("items");
      });
      await context.sync();

      slides.items.forEach((slide) => slide.shapes.items.forEach((shape) => shape.load("id")));
      await context.sync();

      const target = resolveSelectionTarget(
        slides.items.map((slide) => ({
          id: slide.id,
          shapeIds: slide.shapes.items.map((shape) => shape.id),
        })),
        shapeIds
      );
      if (!target) {
        throw new SmartAlignmentError("Those shapes are no longer in the presentation.");
      }

      const slide = slides.items.find((candidate) => candidate.id === target.slideId);
      if (!slide) {
        throw new SmartAlignmentError("Those shapes are no longer in the presentation.");
      }
      slide.setSelectedShapes(target.shapeIds);
      await context.sync();
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not select the shapes.");
  }
}

export interface FallbackTextStyle {
  titleFontName?: string;
  titleFontSize?: number;
  bodyFontName?: string;
  bodyFontSize?: number;
  fontColor?: string;
}

export interface AddSlidesOptions {
  slideSize?: SlideSize;
  fallbackStyle?: FallbackTextStyle;
  stashNotesAsTags?: boolean;
  selectFirstNewSlide?: boolean;
}

interface CatalogRead {
  catalog: LayoutCatalog;
  layoutById: Map<string, PowerPoint.SlideLayout>;
  knownSlideIds: Set<string>;
  slideCount: number;
}

function supports(version: string): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", version);
}

/**
 * Reads every master and its layouts. Three syncs, following the load("items") then per-item
 * load pattern the rest of this file uses rather than nested load paths.
 */
async function readCatalog(context: PowerPoint.RequestContext): Promise<CatalogRead> {
  const masters = context.presentation.slideMasters;
  const slides = context.presentation.slides;
  masters.load("items");
  slides.load("items");
  await context.sync();

  masters.items.forEach((master) => {
    master.load("id,name");
    master.layouts.load("items");
  });
  slides.items.forEach((slide) => slide.load("id"));
  const lastSlide = slides.items[slides.items.length - 1];
  if (lastSlide) lastSlide.slideMaster.load("id");
  await context.sync();

  const layoutProxies: PowerPoint.SlideLayout[] = [];
  masters.items.forEach((master) => {
    master.layouts.items.forEach((layout) => {
      layout.load("id,name");
      layoutProxies.push(layout);
    });
  });
  await context.sync();

  // Layout type is exact and locale-proof, but only exists from 1.8. Its absence is not an error.
  const typed = new Map<string, string>();
  if (supports("1.8")) {
    try {
      const typedProxies = layoutProxies.map((layout) => {
        layout.load("type");
        return layout;
      });
      await context.sync();
      typedProxies.forEach((layout) => typed.set(layout.id, String(layout.type)));
    } catch {
      typed.clear();
    }
  }

  const layouts: LayoutRef[] = [];
  const layoutById = new Map<string, PowerPoint.SlideLayout>();
  masters.items.forEach((master, masterIndex) => {
    master.layouts.items.forEach((layout, layoutIndex) => {
      layouts.push({
        masterId: master.id,
        masterName: master.name,
        masterIndex,
        layoutId: layout.id,
        layoutName: layout.name,
        layoutIndex,
        layoutType: typed.get(layout.id),
      });
      layoutById.set(layout.id, layout);
    });
  });

  return {
    catalog: { layouts, preferredMasterId: lastSlide ? lastSlide.slideMaster.id : undefined },
    layoutById,
    knownSlideIds: new Set(slides.items.map((slide) => slide.id)),
    slideCount: slides.items.length,
  };
}

export async function readLayoutCatalog(): Promise<LayoutCatalog> {
  assertSupported();

  try {
    return await PowerPoint.run(async (context) => (await readCatalog(context)).catalog);
  } catch (error) {
    throw wrapError(error, "PowerPoint could not read the slide layouts.");
  }
}

function toDescriptors(shapes: PowerPoint.Shape[], textCapable: Set<string>): ShapeDescriptor[] {
  return shapes.map((shape) => ({
    id: shape.id,
    name: shape.name,
    type: String(shape.type),
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
    canHoldText: textCapable.has(shape.id),
  }));
}

/** Measures which shapes actually own a text frame, isolating the ones the host refuses. */
async function probeTextFrames(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[]
): Promise<Set<string>> {
  const candidates = shapes.filter((shape) => TEXT_SHAPE_TYPES.has(String(shape.type)));
  const refused = await loadInBatches({
    items: candidates,
    queue: (shape) => shape.textFrame.load("hasText"),
    sync: () => context.sync(),
  });
  const refusedIds = new Set(refused.map((shape) => shape.id));
  return new Set(candidates.filter((shape) => !refusedIds.has(shape.id)).map((shape) => shape.id));
}

interface NewSlide {
  index: number;
  planned: PlannedSlide;
  slide: PowerPoint.Slide;
  layoutId?: string;
}

/**
 * Appends a slide per planned entry using the deck's own master, then fills the layout's title
 * and body placeholders. One bad slide is reported and skipped rather than aborting the run.
 */
export async function addSlidesFromPlan(
  plan: PlannedSlide[],
  options: AddSlidesOptions = {}
): Promise<AddSlidesResult> {
  assertSupported();
  const slides = validatePlan(plan);
  const slideSize = options.slideSize ?? DEFAULT_SLIDE_SIZE;
  const style = options.fallbackStyle ?? {};

  try {
    return await PowerPoint.run(async (context) => {
      const { catalog, layoutById, knownSlideIds, slideCount } = await readCatalog(context);
      const choices = chooseLayouts(catalog, slides);
      const failed: SlideFailure[] = [];

      // Fallback geometry comes from the layout itself, so a drawn box lands where text belongs.
      const usedLayoutIds = Array.from(
        new Set(choices.map((choice) => choice?.layoutId).filter((id): id is string => Boolean(id)))
      );
      const layoutShapes = new Map<string, ShapeDescriptor[]>();
      if (usedLayoutIds.length > 0) {
        usedLayoutIds.forEach((id) => layoutById.get(id)?.shapes.load("items"));
        await context.sync();
        usedLayoutIds.forEach((id) =>
          layoutById.get(id)?.shapes.items.forEach((shape) => queueShapeLoads([shape]))
        );
        await context.sync();
        usedLayoutIds.forEach((id) => {
          const shapes = layoutById.get(id)?.shapes.items ?? [];
          const ids = new Set(shapes.map((shape) => shape.id));
          layoutShapes.set(id, toDescriptors(shapes, ids));
        });
      }

      // Adds are not idempotent, so they are never batched and never bisected.
      const created: NewSlide[] = [];
      for (let index = 0; index < slides.length; index += 1) {
        const planned = slides[index];
        const choice = choices[index];
        try {
          const addOptions: PowerPoint.AddSlideOptions = {};
          if (choice) {
            addOptions.layoutId = choice.layoutId;
            addOptions.slideMasterId = choice.slideMasterId;
          }
          context.presentation.slides.add(addOptions);
          // add() is not idempotent, so it can never share a batch or be replayed by a bisect.
          // eslint-disable-next-line office-addins/no-context-sync-in-loop
          await context.sync();
        } catch {
          failed.push({
            index: index + 1,
            title: planned.title,
            reason: "PowerPoint refused the layout for this slide",
          });
          continue;
        }

        const expected = slideCount + created.length;
        const slide = context.presentation.slides.getItemAt(expected);
        slide.load("id");
        // The file's other ops all use load("items") then per-item loads, which is the
        // pattern proven against real hosts here.
        // eslint-disable-next-line office-addins/no-navigational-load
        slide.shapes.load("items");
        try {
          // One sync per added slide is the cost of locating a slide add() gives no handle to.
          // eslint-disable-next-line office-addins/no-context-sync-in-loop
          await context.sync();
        } catch {
          failed.push({
            index: index + 1,
            title: planned.title,
            reason: "PowerPoint could not return the new slide",
          });
          continue;
        }

        if (knownSlideIds.has(slide.id)) {
          // Never write into a slide we are not certain we created.
          failed.push({
            index: index + 1,
            title: planned.title,
            reason: "the slide list changed while Slideware was building",
          });
          continue;
        }
        knownSlideIds.add(slide.id);
        created.push({ index: index + 1, planned, slide, layoutId: choice?.layoutId });
      }

      if (created.length === 0) {
        return {
          added: 0,
          failed,
          blankSlidesLeft: 0,
          notesSkipped: 0,
          fallbackTextBoxes: 0,
        };
      }

      const allShapes: PowerPoint.Shape[] = [];
      created.forEach((entry) => allShapes.push(...entry.slide.shapes.items));
      allShapes.forEach((shape) => queueShapeLoads([shape]));
      await context.sync();

      const textCapable = await probeTextFrames(context, allShapes);

      interface WritePlan {
        entry: NewSlide;
        titleShape?: PowerPoint.Shape;
        bodyShape?: PowerPoint.Shape;
        body: ReturnType<typeof bodyTextFor>;
      }

      const writes: WritePlan[] = created.map((entry) => {
        const shapes = entry.slide.shapes.items;
        const pick = pickPlaceholders(toDescriptors(shapes, textCapable));
        return {
          entry,
          titleShape: shapes.find((shape) => shape.id === pick.titleShapeId),
          bodyShape: shapes.find((shape) => shape.id === pick.bodyShapeId),
          body: bodyTextFor(entry.planned),
        };
      });

      // Text assignments are idempotent, so a failed batch can be bisected safely.
      const refusedWrites = await loadInBatches({
        items: writes,
        queue: (write) => {
          if (write.titleShape)
            write.titleShape.textFrame.textRange.text = write.entry.planned.title;
          if (write.bodyShape && write.body.text.length > 0) {
            write.bodyShape.textFrame.textRange.text = write.body.text;
          }
        },
        sync: () => context.sync(),
      });
      const refused = new Set(refusedWrites.map((write) => write.entry.index));

      let fallbackTextBoxes = 0;
      let blankSlidesLeft = 0;
      const filled: NewSlide[] = [];

      for (const write of writes) {
        if (refused.has(write.entry.index)) {
          blankSlidesLeft += 1;
          failed.push({
            index: write.entry.index,
            title: write.entry.planned.title,
            reason: "PowerPoint refused the placeholder text, so a blank slide was left",
          });
          continue;
        }

        const needsTitle = !write.titleShape;
        const needsBody = !write.bodyShape && write.body.text.length > 0;
        if (needsTitle || needsBody) {
          // addTextBox cannot be bisected: a replay would duplicate the box. Keep it serial.
          const rects = fallbackRects(
            layoutShapes.get(write.entry.layoutId ?? "") ?? [],
            slideSize
          );
          try {
            if (needsTitle) {
              const box = write.entry.slide.shapes.addTextBox(
                write.entry.planned.title,
                rects.title
              );
              const font = box.textFrame.textRange.font;
              if (style.titleFontName) font.name = style.titleFontName;
              if (style.titleFontSize) font.size = style.titleFontSize;
              if (style.fontColor) font.color = style.fontColor;
              font.bold = true;
            }
            if (needsBody) {
              const box = write.entry.slide.shapes.addTextBox(write.body.text, rects.body);
              const font = box.textFrame.textRange.font;
              if (style.bodyFontName) font.name = style.bodyFontName;
              if (style.bodyFontSize) font.size = style.bodyFontSize;
              if (style.fontColor) font.color = style.fontColor;
            }
            // addTextBox is not idempotent: a bisect replay would duplicate the box.
            // eslint-disable-next-line office-addins/no-context-sync-in-loop
            await context.sync();
            fallbackTextBoxes += 1;
          } catch {
            failed.push({
              index: write.entry.index,
              title: write.entry.planned.title,
              reason: "PowerPoint refused a text box for this slide",
            });
            continue;
          }
        }
        filled.push(write.entry);
      }

      // Cosmetics must never cost a slide its content, so this batch's failure is swallowed.
      let notesSkipped = 0;
      try {
        writes.forEach((write) => {
          if (write.bodyShape && write.body.useBullets && write.body.text.length > 0) {
            write.bodyShape.textFrame.textRange.paragraphFormat.bulletFormat.visible = true;
          }
        });
        if (options.stashNotesAsTags !== false) {
          filled.forEach((entry) => {
            if (entry.planned.notes) {
              entry.slide.tags.add("SlidewareNotes", entry.planned.notes);
              notesSkipped += 1;
            }
          });
        } else {
          notesSkipped = filled.filter((entry) => Boolean(entry.planned.notes)).length;
        }
        if (options.selectFirstNewSlide !== false && filled.length > 0) {
          context.presentation.setSelectedSlides([filled[0].slide.id]);
        }
        await context.sync();
      } catch {
        notesSkipped = filled.filter((entry) => Boolean(entry.planned.notes)).length;
      }

      return {
        added: filled.length,
        failed,
        blankSlidesLeft,
        notesSkipped,
        fallbackTextBoxes,
        firstNewSlideId: filled.length > 0 ? filled[0].slide.id : undefined,
      };
    });
  } catch (error) {
    throw wrapError(error, "PowerPoint could not build the slides.");
  }
}
