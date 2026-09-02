const test = require("node:test");
const assert = require("node:assert/strict");
const { pickPlaceholders, fallbackRects } = require("../lib-test/features/placeholders.js");

function shape(overrides) {
  return {
    id: "s1",
    name: "Shape",
    type: "Placeholder",
    left: 60,
    top: 60,
    width: 400,
    height: 80,
    canHoldText: true,
    ...overrides,
  };
}

const SLIDE = { width: 960, height: 540 };

test("picks a title placeholder by name regardless of array order", () => {
  const pick = pickPlaceholders([
    shape({ id: "sub", name: "Subtitle 2", top: 300 }),
    shape({ id: "title", name: "Title 1", top: 60 }),
  ]);
  assert.equal(pick.titleShapeId, "title");
});

test("prefers the placeholder type when the host reports it", () => {
  const pick = pickPlaceholders([
    shape({ id: "a", name: "Rectangle 4", top: 40, placeholderType: "Title" }),
    shape({ id: "b", name: "Title 1", top: 20 }),
  ]);
  assert.equal(pick.titleShapeId, "a");
});

test("falls back to the topmost text shape when nothing is named title", () => {
  const pick = pickPlaceholders([
    shape({ id: "low", name: "Rectangle 3", top: 300 }),
    shape({ id: "high", name: "Rectangle 2", top: 40 }),
  ]);
  assert.equal(pick.titleShapeId, "high");
});

test("ignores shapes that cannot hold text", () => {
  const pick = pickPlaceholders([
    shape({ id: "pic", name: "Title 1", top: 10, canHoldText: false }),
    shape({ id: "real", name: "Rectangle 2", top: 90 }),
  ]);
  assert.equal(pick.titleShapeId, "real");
});

test("picks the largest body below the title", () => {
  const pick = pickPlaceholders([
    shape({ id: "title", name: "Title 1", top: 40, height: 80 }),
    shape({ id: "small", name: "Content Placeholder 2", top: 150, width: 200, height: 100 }),
    shape({ id: "big", name: "Content Placeholder 3", top: 150, width: 600, height: 300 }),
  ]);
  assert.equal(pick.titleShapeId, "title");
  assert.equal(pick.bodyShapeId, "big");
});

test("never picks a picture or chart placeholder as the body", () => {
  const pick = pickPlaceholders([
    shape({ id: "title", name: "Title 1", top: 40 }),
    shape({ id: "pic", name: "Picture Placeholder 2", top: 150, width: 800, height: 400 }),
  ]);
  assert.equal(pick.bodyShapeId, undefined);
});

test("reports no body when the layout only has a title", () => {
  const pick = pickPlaceholders([shape({ id: "title", name: "Title 1", top: 40 })]);
  assert.equal(pick.titleShapeId, "title");
  assert.equal(pick.bodyShapeId, undefined);
});

test("returns nothing for an empty or text-free shape list", () => {
  assert.deepEqual(pickPlaceholders([]), {});
  assert.deepEqual(pickPlaceholders([shape({ canHoldText: false })]), {});
});

test("is deterministic when geometry ties", () => {
  const shapes = [
    shape({ id: "b", name: "Rectangle 2", top: 40, left: 500 }),
    shape({ id: "a", name: "Rectangle 1", top: 40, left: 100 }),
  ];
  assert.equal(pickPlaceholders(shapes).titleShapeId, "a");
  assert.equal(pickPlaceholders(shapes.slice().reverse()).titleShapeId, "a");
});

test("fallback rectangles follow the layout geometry when it exists", () => {
  const rects = fallbackRects(
    [
      shape({ id: "t", name: "Title 1", left: 40, top: 30, width: 700, height: 90 }),
      shape({ id: "b", name: "Content Placeholder 2", left: 40, top: 150, width: 700, height: 300 }),
    ],
    SLIDE
  );
  assert.deepEqual(rects.title, { left: 40, top: 30, width: 700, height: 90 });
  assert.deepEqual(rects.body, { left: 40, top: 150, width: 700, height: 300 });
});

test("fallback rectangles derive from the slide when the layout gives nothing", () => {
  const rects = fallbackRects([], SLIDE);
  assert.equal(rects.title.left, 60);
  assert.ok(rects.title.width <= SLIDE.width - 120);
  assert.ok(rects.body.top > rects.title.top);
  assert.ok(rects.body.top + rects.body.height <= SLIDE.height);
});
