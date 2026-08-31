const test = require("node:test");
const assert = require("node:assert/strict");
const { audit } = require("../lib-test/features/checker.js");

function shape(overrides) {
  return {
    id: "s1",
    name: "Shape 1",
    type: "GeometricShape",
    left: 100,
    top: 100,
    width: 200,
    height: 100,
    text: "Fine text",
    fontName: "Segoe UI",
    fontSize: 18,
    ...overrides,
  };
}

function deck(slides) {
  return { slideCount: slides.length, slides };
}

function slide(index, shapes) {
  return { id: `slide-${index}`, index, shapes };
}

test("flags shapes outside the slide bounds", () => {
  const result = audit(deck([slide(1, [shape({ id: "off", left: 900, width: 200 })])]));
  const finding = result.find((item) => item.rule === "off-slide");
  assert.ok(finding);
  assert.equal(finding.slideIndex, 1);
  assert.equal(finding.shapeId, "off");
});

test("accepts shapes fully inside the slide", () => {
  const result = audit(deck([slide(1, [shape()])]));
  assert.equal(result.find((item) => item.rule === "off-slide"), undefined);
});

test("flags fonts smaller than 12 points", () => {
  const result = audit(deck([slide(1, [shape({ fontSize: 9 })])]));
  assert.ok(result.find((item) => item.rule === "tiny-font"));
  const clean = audit(deck([slide(1, [shape({ fontSize: 12 })])]));
  assert.equal(clean.find((item) => item.rule === "tiny-font"), undefined);
});

test("flags decks using more than three font families", () => {
  const shapes = ["Arial", "Georgia", "Verdana", "Courier New"].map((fontName, index) =>
    shape({ id: `f${index}`, fontName })
  );
  const result = audit(deck([slide(1, shapes)]));
  const finding = result.find((item) => item.rule === "font-sprawl");
  assert.ok(finding);
  assert.equal(finding.shapeId, undefined);
  const clean = audit(deck([slide(1, shapes.slice(0, 3))]));
  assert.equal(clean.find((item) => item.rule === "font-sprawl"), undefined);
});

test("flags empty text boxes but not empty pictures", () => {
  const result = audit(
    deck([slide(1, [shape({ id: "empty", text: "  " }), shape({ id: "pic", type: "Image", text: "" })])])
  );
  const findings = result.filter((item) => item.rule === "empty-text");
  assert.deepEqual(findings.map((item) => item.shapeId), ["empty"]);
});

test("flags overlong text blocks past 300 characters", () => {
  const result = audit(deck([slide(2, [shape({ id: "long", text: "x".repeat(301) })])]));
  const finding = result.find((item) => item.rule === "overlong");
  assert.ok(finding);
  assert.equal(finding.slideIndex, 2);
});
