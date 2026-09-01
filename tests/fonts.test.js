const test = require("node:test");
const assert = require("node:assert/strict");
const { usedFonts, fontReplaceFormats } = require("../lib-test/features/fonts.js");

function shape(overrides) {
  return {
    id: "a",
    name: "TextBox",
    type: "TextBox",
    left: 0,
    top: 0,
    width: 10,
    height: 10,
    text: "Hello",
    fontName: "Arial",
    ...overrides,
  };
}

function deck(shapes) {
  return { slideCount: 1, slides: [{ id: "s1", index: 1, shapes }] };
}

test("counts fonts used by text shapes, most used first", () => {
  const result = usedFonts(
    deck([
      shape({ id: "a", fontName: "Georgia" }),
      shape({ id: "b", fontName: "Arial" }),
      shape({ id: "c", fontName: "Georgia" }),
      shape({ id: "d", text: "", fontName: "Courier" }),
    ])
  );
  assert.deepEqual(result, [
    { name: "Georgia", count: 2 },
    { name: "Arial", count: 1 },
  ]);
});

test("builds replacement formats for matching fonts case-insensitively", () => {
  const formats = fontReplaceFormats(
    deck([
      shape({ id: "a", fontName: "arial" }),
      shape({ id: "b", fontName: "Georgia" }),
      shape({ id: "c", fontName: "Arial" }),
    ]),
    "ARIAL",
    "Inter"
  );
  assert.deepEqual(formats, [
    { id: "a", fontName: "Inter" },
    { id: "c", fontName: "Inter" },
  ]);
});

test("rejects blank fonts and reports when nothing matches", () => {
  assert.throws(() => fontReplaceFormats(deck([shape()]), " ", "Inter"), /which font to replace/);
  assert.throws(() => fontReplaceFormats(deck([shape()]), "Arial", ""), /replacement font/);
  assert.throws(() => fontReplaceFormats(deck([shape()]), "Comic Sans MS", "Inter"), /No shapes use/);
});
