const test = require("node:test");
const assert = require("node:assert/strict");
const { normalizeLayoutName, chooseLayout, chooseLayouts } = require("../lib-test/features/slidelayout.js");

function layout(masterIndex, masterId, layoutIndex, layoutName, layoutType) {
  return {
    masterId,
    masterName: `Master ${masterIndex}`,
    masterIndex,
    layoutId: `${masterId}-L${layoutIndex}`,
    layoutName,
    layoutIndex,
    layoutType,
  };
}

const NAMED = {
  layouts: [
    layout(0, "m0", 0, "Title Slide"),
    layout(0, "m0", 1, "Title and Content"),
    layout(0, "m0", 2, "Section Header"),
  ],
  preferredMasterId: "m0",
};

test("normalizeLayoutName folds case, punctuation, and separators", () => {
  assert.equal(normalizeLayoutName("Title and Content"), "title and content");
  assert.equal(normalizeLayoutName("TITLE_AND_CONTENT"), "title and content");
  assert.equal(normalizeLayoutName("Title & Content"), "title content");
  assert.equal(normalizeLayoutName("  Section   Header  "), "section header");
});

test("matches by layout name when no type is available", () => {
  assert.equal(chooseLayout(NAMED, "title").layoutName, "Title Slide");
  assert.equal(chooseLayout(NAMED, "bullets").layoutName, "Title and Content");
  assert.equal(chooseLayout(NAMED, "section").layoutName, "Section Header");
  assert.equal(chooseLayout(NAMED, "bullets").matchedBy, "name");
});

test("prefers the layout type when the host reports it", () => {
  const catalog = {
    layouts: [
      layout(0, "m0", 0, "Diapositive de titre", "Title"),
      layout(0, "m0", 1, "Titre et contenu", "Object"),
      layout(0, "m0", 2, "Titre de section", "SectionHeader"),
    ],
    preferredMasterId: "m0",
  };
  assert.equal(chooseLayout(catalog, "title").layoutName, "Diapositive de titre");
  assert.equal(chooseLayout(catalog, "bullets").layoutName, "Titre et contenu");
  assert.equal(chooseLayout(catalog, "section").layoutName, "Titre de section");
  assert.equal(chooseLayout(catalog, "title").matchedBy, "type");
});

test("scopes to the preferred master and never pairs a layout with another master", () => {
  const catalog = {
    layouts: [
      layout(0, "m0", 0, "Title Slide"),
      layout(0, "m0", 1, "Title and Content"),
      layout(1, "m1", 0, "Title Slide"),
      layout(1, "m1", 1, "Title and Content"),
    ],
    preferredMasterId: "m1",
  };
  const choice = chooseLayout(catalog, "bullets");
  assert.equal(choice.slideMasterId, "m1");
  assert.equal(choice.layoutId, "m1-L1");
});

test("falls back to the first master when the preferred one is unknown", () => {
  const catalog = { layouts: NAMED.layouts, preferredMasterId: "missing" };
  assert.equal(chooseLayout(catalog, "title").slideMasterId, "m0");
});

test("falls back to layout position when nothing matches by name", () => {
  const catalog = {
    layouts: [layout(0, "m0", 0, "Alpha"), layout(0, "m0", 1, "Beta")],
    preferredMasterId: "m0",
  };
  assert.equal(chooseLayout(catalog, "title").layoutName, "Alpha");
  assert.equal(chooseLayout(catalog, "bullets").layoutName, "Beta");
  assert.equal(chooseLayout(catalog, "title").matchedBy, "index");
});

test("clamps the position fallback to what the master actually has", () => {
  const catalog = { layouts: [layout(0, "m0", 0, "Only")], preferredMasterId: "m0" };
  assert.equal(chooseLayout(catalog, "bullets").layoutName, "Only");
});

test("returns undefined when there are no layouts at all", () => {
  assert.equal(chooseLayout({ layouts: [] }, "bullets"), undefined);
});

test("chooseLayouts maps a whole plan in order", () => {
  const plan = [
    { kind: "title", title: "A", bullets: [] },
    { kind: "bullets", title: "B", bullets: ["x"] },
  ];
  assert.deepEqual(
    chooseLayouts(NAMED, plan).map((choice) => choice.layoutName),
    ["Title Slide", "Title and Content"]
  );
});
