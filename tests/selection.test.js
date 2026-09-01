const test = require("node:test");
const assert = require("node:assert/strict");
const { matchShapes } = require("../lib-test/features/selection.js");

function shape(overrides) {
  return {
    id: "x",
    name: "Shape",
    type: "GeometricShape",
    left: 0,
    top: 0,
    width: 100,
    height: 50,
    text: "",
    ...overrides,
  };
}

const anchor = shape({ id: "anchor", fillColor: "#FF0000" });

test("matches shapes of the same type", () => {
  const all = [
    anchor,
    shape({ id: "b", type: "GeometricShape", width: 10 }),
    shape({ id: "c", type: "Image" }),
  ];
  assert.deepEqual(matchShapes(all, anchor, { sameType: true }), ["anchor", "b"]);
});

test("matches fill colors case-insensitively", () => {
  const all = [
    anchor,
    shape({ id: "b", fillColor: "#ff0000" }),
    shape({ id: "c", fillColor: "#00FF00" }),
    shape({ id: "d" }),
  ];
  assert.deepEqual(matchShapes(all, anchor, { sameFill: true }), ["anchor", "b"]);
});

test("matches size within one point of tolerance", () => {
  const all = [
    anchor,
    shape({ id: "b", width: 100.9, height: 49.2 }),
    shape({ id: "c", width: 102, height: 50 }),
  ];
  assert.deepEqual(matchShapes(all, anchor, { sameSize: true }), ["anchor", "b"]);
});

test("intersects combined criteria", () => {
  const all = [
    anchor,
    shape({ id: "b", fillColor: "#FF0000", width: 100 }),
    shape({ id: "c", fillColor: "#FF0000", width: 300 }),
    shape({ id: "d", fillColor: "#0000FF", width: 100 }),
  ];
  assert.deepEqual(matchShapes(all, anchor, { sameFill: true, sameSize: true }), ["anchor", "b"]);
});

test("rejects a call with no criteria", () => {
  assert.throws(() => matchShapes([anchor], anchor, {}), /Pick at least one criteria/);
});

test("matches by font family, font color, and font size", () => {
  const styledAnchor = shape({
    id: "anchor",
    fontName: "Georgia",
    fontColor: "#112233",
    fontSize: 20,
  });
  const all = [
    styledAnchor,
    shape({ id: "b", fontName: "georgia", fontColor: "#112233", fontSize: 20 }),
    shape({ id: "c", fontName: "Arial", fontColor: "#112233", fontSize: 20 }),
    shape({ id: "d", fontName: "Georgia", fontColor: "#FF0000", fontSize: 20 }),
    shape({ id: "e", fontName: "Georgia", fontColor: "#112233", fontSize: 12 }),
  ];
  assert.deepEqual(
    matchShapes(all, styledAnchor, { sameFont: true, sameFontColor: true, sameFontSize: true }),
    ["anchor", "b"]
  );
  assert.deepEqual(matchShapes(all, styledAnchor, { sameFont: true }), ["anchor", "b", "d", "e"]);
});

const { resolveSelectionTarget } = require("../lib-test/features/selection.js");

const DECK = [
  { id: "slide-1", shapeIds: ["a", "b", "c"] },
  { id: "slide-2", shapeIds: ["d", "e"] },
];

test("targets the slide that owns the requested shapes", () => {
  assert.deepEqual(resolveSelectionTarget(DECK, ["d", "e"]), {
    slideId: "slide-2",
    shapeIds: ["d", "e"],
  });
});

test("keeps the requested order of the shape ids", () => {
  assert.deepEqual(resolveSelectionTarget(DECK, ["c", "a"]), {
    slideId: "slide-1",
    shapeIds: ["c", "a"],
  });
});

test("drops ids that live on another slide because selection is per slide", () => {
  assert.deepEqual(resolveSelectionTarget(DECK, ["a", "b", "d"]), {
    slideId: "slide-1",
    shapeIds: ["a", "b"],
  });
});

test("picks the slide holding the most requested shapes", () => {
  assert.deepEqual(resolveSelectionTarget(DECK, ["a", "d", "e"]), {
    slideId: "slide-2",
    shapeIds: ["d", "e"],
  });
});

test("prefers the earlier slide when the counts tie", () => {
  assert.deepEqual(resolveSelectionTarget(DECK, ["a", "d"]), {
    slideId: "slide-1",
    shapeIds: ["a"],
  });
});

test("returns undefined when no requested shape is still in the deck", () => {
  assert.equal(resolveSelectionTarget(DECK, ["gone"]), undefined);
  assert.equal(resolveSelectionTarget(DECK, []), undefined);
});
