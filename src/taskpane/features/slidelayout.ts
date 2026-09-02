import { PlannedSlide, SlideKind } from "./deck-plan";

export interface LayoutRef {
  masterId: string;
  masterName: string;
  masterIndex: number;
  layoutId: string;
  layoutName: string;
  layoutIndex: number;
  /** Only populated when the host is PowerPointApi 1.8 or newer. */
  layoutType?: string;
}

export interface LayoutCatalog {
  layouts: LayoutRef[];
  /** Master of the deck's last slide, so new slides inherit the look already in use. */
  preferredMasterId?: string;
}

export interface LayoutChoice {
  slideMasterId: string;
  layoutId: string;
  layoutName: string;
  matchedBy: "type" | "name" | "index";
}

const TYPES: Record<SlideKind, string[]> = {
  title: ["Title"],
  section: ["SectionHeader"],
  bullets: ["Object", "Text", "TitleOnly", "Custom"],
};

const NAMES: Record<SlideKind, string[]> = {
  title: ["title slide", "title", "cover", "opening"],
  section: ["section header", "section", "divider", "chapter"],
  bullets: ["title and content", "title content", "content", "bullets", "text"],
};

const INDEX_FALLBACK: Record<SlideKind, number> = { title: 0, section: 1, bullets: 1 };

export function normalizeLayoutName(name: string): string {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function scopeToMaster(catalog: LayoutCatalog): LayoutRef[] {
  const preferred = catalog.preferredMasterId
    ? catalog.layouts.filter((layout) => layout.masterId === catalog.preferredMasterId)
    : [];
  if (preferred.length > 0) return preferred;

  const first = catalog.layouts.filter((layout) => layout.masterIndex === 0);
  return first.length > 0 ? first : catalog.layouts;
}

function toChoice(layout: LayoutRef, matchedBy: LayoutChoice["matchedBy"]): LayoutChoice {
  // The master always travels with the layout: a layoutId from another master is refused by the host.
  return {
    slideMasterId: layout.masterId,
    layoutId: layout.layoutId,
    layoutName: layout.layoutName,
    matchedBy,
  };
}

export function chooseLayout(catalog: LayoutCatalog, kind: SlideKind): LayoutChoice | undefined {
  const scoped = scopeToMaster(catalog);
  if (scoped.length === 0) return undefined;

  for (const wanted of TYPES[kind]) {
    const hit = scoped.find((layout) => layout.layoutType === wanted);
    if (hit) return toChoice(hit, "type");
  }

  const named = scoped.map((layout) => ({ layout, key: normalizeLayoutName(layout.layoutName) }));
  for (const wanted of NAMES[kind]) {
    const exact = named.find((entry) => entry.key === wanted);
    if (exact) return toChoice(exact.layout, "name");
  }
  for (const wanted of NAMES[kind]) {
    const prefix = named.find((entry) => entry.key.startsWith(wanted));
    if (prefix) return toChoice(prefix.layout, "name");
  }
  for (const wanted of NAMES[kind]) {
    const loose = named.find((entry) => entry.key.includes(wanted));
    if (loose) return toChoice(loose.layout, "name");
  }

  const ordered = scoped.slice().sort((a, b) => a.layoutIndex - b.layoutIndex);
  const position = Math.min(INDEX_FALLBACK[kind], ordered.length - 1);
  return toChoice(ordered[position], "index");
}

export function chooseLayouts(
  catalog: LayoutCatalog,
  plan: PlannedSlide[]
): (LayoutChoice | undefined)[] {
  return plan.map((slide) => chooseLayout(catalog, slide.kind));
}
