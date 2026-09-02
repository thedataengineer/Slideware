import { cleanString, cleanStringList, extractJson } from "./model-json";

export const MAX_SLIDES = 40;
export const MAX_BULLETS = 6;
export const MAX_TITLE_CHARS = 120;
export const MAX_BULLET_CHARS = 200;
export const MAX_NOTES_CHARS = 600;

const UNREADABLE =
  "The model's slide plan could not be read as JSON. Try again, or pick a stronger model in the Gen AI tab.";
const NO_SLIDES =
  "The model returned a plan with no usable slides. Try again, or pick a stronger model in the Gen AI tab.";

export type SlideKind = "title" | "section" | "bullets";

export interface PlannedSlide {
  kind: SlideKind;
  title: string;
  bullets: string[];
  /** Carried and shown in the plan editor. PowerPoint has no notes API, so it is stashed as a tag. */
  notes?: string;
}

export interface SlidePlan {
  slides: PlannedSlide[];
}

const KIND_SYNONYMS: Record<string, SlideKind> = {
  title: "title",
  title_slide: "title",
  titleslide: "title",
  cover: "title",
  opening: "title",
  section: "section",
  section_header: "section",
  sectionheader: "section",
  divider: "section",
  chapter: "section",
  bullets: "bullets",
  bullet: "bullets",
  content: "bullets",
  body: "bullets",
  points: "bullets",
};

function stripGlyph(bullet: string): string {
  return bullet.replace(/^(?:[-*•‣–—]|\d+[.)])\s*/, "").trim();
}

function toBullets(raw: unknown): string[] {
  return cleanStringList(raw, MAX_BULLETS * 2, MAX_BULLET_CHARS)
    .map(stripGlyph)
    .filter((bullet) => bullet.length > 0)
    .slice(0, MAX_BULLETS);
}

function toKind(raw: unknown, hasBullets: boolean): SlideKind {
  const key = cleanString(raw, 40)
    .toLowerCase()
    .replace(/[\s-]+/g, "_");
  const mapped = KIND_SYNONYMS[key];
  if (mapped) return mapped;
  return hasBullets ? "bullets" : "section";
}

function toSlide(raw: unknown): PlannedSlide | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const source = raw as { kind?: unknown; title?: unknown; bullets?: unknown; notes?: unknown };

  const title = cleanString(source.title, MAX_TITLE_CHARS);
  if (title.length === 0) return undefined;

  const bullets = toBullets(source.bullets);
  // A heading with no bullets is still real content, so keep it as a divider.
  const kind = toKind(source.kind, bullets.length > 0);
  const resolved: SlideKind = kind === "bullets" && bullets.length === 0 ? "section" : kind;

  const slide: PlannedSlide = { kind: resolved, title, bullets };
  const notes = cleanString(source.notes, MAX_NOTES_CHARS);
  if (notes.length > 0) slide.notes = notes;
  return slide;
}

function slidesArray(root: unknown): unknown[] | undefined {
  if (Array.isArray(root)) return root;
  if (root && typeof root === "object") {
    const container = root as { slides?: unknown; plan?: unknown };
    if (Array.isArray(container.slides)) return container.slides;
    if (Array.isArray(container.plan)) return container.plan;
  }
  return undefined;
}

/**
 * Throws when the model gave nothing usable, because there is no honest fallback deck.
 * Salvages individually bad slides the same way the interview parser does.
 */
export function parseSlidePlan(raw: string): SlidePlan {
  const rawSlides = slidesArray(extractJson(raw));
  if (rawSlides === undefined) throw new Error(UNREADABLE);

  const slides: PlannedSlide[] = [];
  rawSlides.forEach((entry) => {
    if (slides.length >= MAX_SLIDES) return;
    const slide = toSlide(entry);
    if (slide) slides.push(slide);
  });

  if (slides.length === 0) throw new Error(NO_SLIDES);
  return { slides };
}

function assertIndex(plan: SlidePlan, index: number): void {
  if (!Number.isInteger(index) || index < 0 || index >= plan.slides.length) {
    throw new Error(`There is no slide ${index + 1} in the plan.`);
  }
}

export function movePlanSlide(plan: SlidePlan, index: number, delta: number): SlidePlan {
  assertIndex(plan, index);
  const target = index + delta;
  if (target < 0 || target >= plan.slides.length) return plan;

  const slides = plan.slides.slice();
  const [moved] = slides.splice(index, 1);
  slides.splice(target, 0, moved);
  return { slides };
}

export function removePlanSlide(plan: SlidePlan, index: number): SlidePlan {
  assertIndex(plan, index);
  return { slides: plan.slides.filter((_, position) => position !== index) };
}

export function updatePlanSlide(
  plan: SlidePlan,
  index: number,
  patch: Partial<PlannedSlide>
): SlidePlan {
  assertIndex(plan, index);

  const slides = plan.slides.map((slide, position) => {
    if (position !== index) return slide;

    const bullets = patch.bullets === undefined ? slide.bullets : toBullets(patch.bullets);
    const title =
      patch.title === undefined ? slide.title : cleanString(patch.title, MAX_TITLE_CHARS);
    const kind = patch.kind === undefined ? slide.kind : patch.kind;
    const next: PlannedSlide = {
      kind: kind === "bullets" && bullets.length === 0 ? "section" : kind,
      title: title.length > 0 ? title : slide.title,
      bullets,
    };
    const notes =
      patch.notes === undefined ? slide.notes : cleanString(patch.notes, MAX_NOTES_CHARS);
    if (notes && notes.length > 0) next.notes = notes;
    return next;
  });

  return { slides };
}

function plural(count: number, noun: string): string {
  return `${count} ${noun}${count === 1 ? "" : "s"}`;
}

export function planSummary(plan: SlidePlan): string {
  const titles = plan.slides.filter((slide) => slide.kind === "title").length;
  const sections = plan.slides.filter((slide) => slide.kind === "section").length;
  const content = plan.slides.filter((slide) => slide.kind === "bullets").length;
  return `${plural(plan.slides.length, "slide")} · ${titles} title, ${plural(sections, "section")}, ${content} content`;
}

export function slideMeta(slide: PlannedSlide): string {
  const base =
    slide.kind === "title"
      ? "Title slide"
      : slide.kind === "section"
        ? "Section divider"
        : `Bullets · ${plural(slide.bullets.length, "bullet")}`;
  return slide.notes ? `${base} · has notes` : base;
}

export interface SlideFailure {
  /** 1-based position in the submitted plan, so the message matches what the user sees. */
  index: number;
  title: string;
  reason: string;
}

export interface AddSlidesResult {
  /** Slides that were created and filled. A created-but-empty slide counts as a failure. */
  added: number;
  failed: SlideFailure[];
  blankSlidesLeft: number;
  notesSkipped: number;
  fallbackTextBoxes: number;
  firstNewSlideId?: string;
}

export interface BodyText {
  text: string;
  /** Ask the host for bullet glyphs rather than typing them, so a placeholder cannot double them. */
  useBullets: boolean;
}

export function validatePlan(plan: PlannedSlide[]): PlannedSlide[] {
  const cleaned = plan
    .map((slide) => {
      const title = cleanString(slide.title, MAX_TITLE_CHARS);
      return title.length === 0 ? undefined : { ...slide, title };
    })
    .filter((slide): slide is PlannedSlide => slide !== undefined);

  if (cleaned.length === 0) throw new Error("There are no slides to build.");
  if (cleaned.length > MAX_SLIDES) {
    throw new Error(`Slideware builds up to ${MAX_SLIDES} slides at a time; trim the plan.`);
  }
  return cleaned;
}

export function bodyTextFor(slide: PlannedSlide): BodyText {
  if (slide.kind === "section") return { text: "", useBullets: false };
  if (slide.kind === "title") return { text: slide.bullets.join("\n"), useBullets: false };
  return { text: slide.bullets.join("\n"), useBullets: true };
}

export function summarizeResult(result: AddSlidesResult): string {
  const parts = [`Added ${plural(result.added, "slide")}.`];

  if (result.failed.length > 0) {
    const detail = result.failed
      .map((failure) => `slide ${failure.index}: ${failure.reason}`)
      .join("; ");
    parts.push(`${plural(result.failed.length, "slide")} could not be built (${detail}).`);
  }
  if (result.blankSlidesLeft > 0) {
    parts.push(`${plural(result.blankSlidesLeft, "blank slide")} left in the deck.`);
  }
  if (result.notesSkipped > 0) {
    parts.push(
      `PowerPoint has no notes API, so ${plural(result.notesSkipped, "note")} were kept on the slides as tags.`
    );
  }
  return parts.join(" ");
}
