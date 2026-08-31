const test = require("node:test");
const assert = require("node:assert/strict");
const { templateShapes } = require("../lib-test/features/templates.js");
const { defaultBrand } = require("../lib-test/features/branding.js");

const brand = defaultBrand();
const slideSize = { width: 960, height: 540 };

function assertWithinBounds(specs) {
  specs.forEach((spec) => {
    assert.ok(spec.left >= 0 && spec.top >= 0, "spec starts on the slide");
    assert.ok(spec.left + spec.width <= slideSize.width, "spec fits horizontally");
    assert.ok(spec.top + spec.height <= slideSize.height, "spec fits vertically");
  });
}

test("title block emits heading and subtitle in brand fonts", () => {
  const specs = templateShapes("title-block", brand, slideSize);
  assert.equal(specs.length, 2);
  assert.equal(specs[0].fontName, brand.headingFont);
  assert.equal(specs[1].fontName, brand.bodyFont);
  assertWithinBounds(specs);
});

test("kpi row emits three tiles with labels", () => {
  const specs = templateShapes("kpi-row", brand, slideSize);
  const rects = specs.filter((spec) => spec.kind === "rect");
  assert.equal(rects.length, 3);
  assertWithinBounds(specs);
});

test("quote card and section divider fit the slide", () => {
  assertWithinBounds(templateShapes("quote-card", brand, slideSize));
  assertWithinBounds(templateShapes("section-divider", brand, slideSize));
});

test("rejects unknown template names", () => {
  assert.throws(() => templateShapes("nope", brand, slideSize), /Unknown template/);
});
