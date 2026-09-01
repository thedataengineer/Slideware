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
