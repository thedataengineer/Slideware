const test = require("node:test");
const assert = require("node:assert/strict");
const {
  defaultBrand,
  parseBrand,
  serializeBrand,
  normalizeHex,
  brandSelectionFormats,
  brandDeckFormats,
} = require("../lib-test/features/branding.js");

test("round-trips a brand through serialization", () => {
  const brand = {
    headingFont: "Georgia",
    bodyFont: "Verdana",
    colors: ["#111111", "#222222", "#333333", "#444444", "#555555", "#666666"],
  };
  assert.deepEqual(parseBrand(serializeBrand(brand)), brand);
});

test("falls back to the default brand on corrupt input", () => {
  assert.deepEqual(parseBrand("not json"), defaultBrand());
  assert.deepEqual(parseBrand(JSON.stringify({ headingFont: 3 })), defaultBrand());
  assert.deepEqual(parseBrand(null), defaultBrand());
});

test("normalizes three-digit hex and validates input", () => {
  assert.equal(normalizeHex("#abc"), "#aabbcc");
  assert.equal(normalizeHex("#A1B2C3"), "#a1b2c3");
  assert.throws(() => normalizeHex("red"), /valid hex color/);
});

function shape(overrides) {
  return {
    id: "s1",
    name: "Shape",
    type: "GeometricShape",
    left: 0,
    top: 0,
    width: 10,
    height: 10,
    text: "Hello",
    ...overrides,
  };
}

test("builds selection formats only for text-bearing shapes", () => {
  const brand = defaultBrand();
  const formats = brandSelectionFormats(
    [shape({ id: "a" }), shape({ id: "b", text: "" }), shape({ id: "c", type: "Image", text: "x" })],
    brand
  );
  assert.deepEqual(formats, [{ id: "a", fontName: brand.bodyFont, fontColor: brand.colors[0] }]);
});

test("builds deck-wide font fixes for every text shape", () => {
  const brand = defaultBrand();
  const deck = {
    slideCount: 2,
    slides: [
      { id: "s1", index: 1, shapes: [shape({ id: "a" }), shape({ id: "b", text: " " })] },
      { id: "s2", index: 2, shapes: [shape({ id: "c" })] },
    ],
  };
  assert.deepEqual(brandDeckFormats(deck, brand), [
    { id: "a", fontName: brand.bodyFont },
    { id: "c", fontName: brand.bodyFont },
  ]);
});
