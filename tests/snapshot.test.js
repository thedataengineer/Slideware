const test = require("node:test");
const assert = require("node:assert/strict");
const { deriveTitle } = require("../lib-test/features/snapshot.js");

test("prefers a shape whose name contains title", () => {
  const shapes = [
    { id: "a", name: "Content 1", top: 10, text: "Body text" },
    { id: "b", name: "Title 1", top: 200, text: "Quarterly Review" },
  ];
  assert.equal(deriveTitle(shapes), "Quarterly Review");
});

test("falls back to the topmost text shape", () => {
  const shapes = [
    { id: "a", name: "Content 2", top: 150, text: "Lower text" },
    { id: "b", name: "Content 1", top: 40, text: "Upper text" },
    { id: "c", name: "Picture 1", top: 5, text: "" },
  ];
  assert.equal(deriveTitle(shapes), "Upper text");
});

test("returns undefined when no shape has text", () => {
  const shapes = [{ id: "a", name: "Picture 1", top: 5, text: "" }];
  assert.equal(deriveTitle(shapes), undefined);
});

test("ignores title-named shapes without text", () => {
  const shapes = [
    { id: "a", name: "Title 1", top: 10, text: "  " },
    { id: "b", name: "Content 1", top: 90, text: "Real heading" },
  ];
  assert.equal(deriveTitle(shapes), "Real heading");
});

test("does not mistake a subtitle for the slide title", () => {
  const shapes = [
    { id: "a", name: "Subtitle 2", top: 300, text: "A tagline" },
    { id: "b", name: "Title 1", top: 60, text: "Money Moxie" },
  ];
  assert.equal(deriveTitle(shapes), "Money Moxie");
});

test("prefers the topmost title when several shapes are named title", () => {
  const shapes = [
    { id: "a", name: "Title Placeholder 3", top: 400, text: "Lower" },
    { id: "b", name: "Title 1", top: 40, text: "Upper" },
  ];
  assert.equal(deriveTitle(shapes), "Upper");
});
