const test = require("node:test");
const assert = require("node:assert/strict");
const { matchSizes, swapPositions } = require("../lib-test/features/smartbar.js");

const shapes = [
  { id: "ref", left: 10, top: 20, width: 100, height: 50 },
  { id: "b", left: 200, top: 60, width: 40, height: 80 },
  { id: "c", left: 300, top: 90, width: 70, height: 30 },
];

test("matches widths to the first selected shape", () => {
  assert.deepEqual(matchSizes(shapes, "width"), [
    { id: "b", width: 100 },
    { id: "c", width: 100 },
  ]);
});

test("matches heights to the first selected shape", () => {
  assert.deepEqual(matchSizes(shapes, "height"), [
    { id: "b", height: 50 },
    { id: "c", height: 50 },
  ]);
});

test("matches both dimensions to the first selected shape", () => {
  assert.deepEqual(matchSizes(shapes, "both"), [
    { id: "b", width: 100, height: 50 },
    { id: "c", width: 100, height: 50 },
  ]);
});

test("rejects size matching with fewer than two shapes", () => {
  assert.throws(() => matchSizes(shapes.slice(0, 1), "width"), /Select at least 2 shapes/);
});

test("swaps positions of exactly two shapes without touching size", () => {
  assert.deepEqual(swapPositions(shapes.slice(0, 2)), [
    { id: "ref", left: 200, top: 60 },
    { id: "b", left: 10, top: 20 },
  ]);
});

test("rejects swap unless exactly two shapes are selected", () => {
  assert.throws(() => swapPositions(shapes), /Select exactly 2 shapes/);
  assert.throws(() => swapPositions(shapes.slice(0, 1)), /Select exactly 2 shapes/);
});
