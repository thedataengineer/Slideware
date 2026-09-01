const test = require("node:test");
const assert = require("node:assert/strict");
const { splitTextBox, mergeTextBoxes } = require("../lib-test/features/textboxes.js");

function shape(overrides) {
  return {
    id: "a",
    name: "TextBox 1",
    type: "TextBox",
    left: 100,
    top: 50,
    width: 300,
    height: 90,
    text: "Alpha\nBeta\nGamma",
    fontName: "Georgia",
    fontSize: 20,
    fontColor: "#112233",
    ...overrides,
  };
}

test("splits a multi-line text box into stacked per-line boxes", () => {
  const result = splitTextBox(shape());
  assert.deepEqual(result.deleteIds, ["a"]);
  assert.equal(result.inserts.length, 3);
  assert.deepEqual(result.inserts.map((spec) => spec.text), ["Alpha", "Beta", "Gamma"]);
  result.inserts.forEach((spec, index) => {
    assert.equal(spec.kind, "textbox");
    assert.equal(spec.left, 100);
    assert.equal(spec.width, 300);
    assert.equal(spec.height, 30);
    assert.equal(spec.top, 50 + index * 30);
    assert.equal(spec.fontName, "Georgia");
    assert.equal(spec.fontSize, 20);
    assert.equal(spec.fontColor, "#112233");
  });
});

test("split skips blank lines and trims line ends", () => {
  const result = splitTextBox(shape({ text: "One\r\n\r\n  \nTwo" }));
  assert.deepEqual(result.inserts.map((spec) => spec.text), ["One", "Two"]);
  assert.equal(result.inserts[0].height, 45);
});

test("split rejects single-line and empty boxes", () => {
  assert.throws(() => splitTextBox(shape({ text: "Only line" })), /more than one line/);
  assert.throws(() => splitTextBox(shape({ text: "  " })), /more than one line/);
});

test("merges text boxes in selection order with bounding-box geometry", () => {
  const first = shape({ id: "b", left: 400, top: 200, width: 100, height: 40, text: "Second area" });
  const second = shape({ id: "c", left: 100, top: 50, width: 200, height: 60, text: "First area" });
  const result = mergeTextBoxes([first, second]);
  assert.deepEqual(result.deleteIds, ["b", "c"]);
  assert.equal(result.insert.text, "Second area\nFirst area");
  assert.equal(result.insert.left, 100);
  assert.equal(result.insert.top, 50);
  assert.equal(result.insert.width, 400);
  assert.equal(result.insert.height, 190);
  assert.equal(result.insert.fontName, "Georgia");
});

test("merge ignores shapes without text and needs at least two", () => {
  const empty = shape({ id: "e", text: " " });
  const only = shape({ id: "o" });
  assert.throws(() => mergeTextBoxes([empty, only]), /at least 2 text boxes/);
  assert.throws(() => mergeTextBoxes([only]), /at least 2 text boxes/);
});
