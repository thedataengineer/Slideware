const test = require("node:test");
const assert = require("node:assert/strict");
const { buildAgenda } = require("../lib-test/features/agenda.js");

test("numbers agenda lines from slide titles", () => {
  const agenda = buildAgenda(["Market Overview", "Product Roadmap", "Financials"]);
  assert.deepEqual(agenda.lines, ["1. Market Overview", "2. Product Roadmap", "3. Financials"]);
  assert.equal(agenda.text, "1. Market Overview\n2. Product Roadmap\n3. Financials");
});

test("skips empty and whitespace titles", () => {
  const agenda = buildAgenda(["Intro", "", "   ", "Close"]);
  assert.deepEqual(agenda.lines, ["1. Intro", "2. Close"]);
});

test("throws when no titles remain", () => {
  assert.throws(() => buildAgenda(["", undefined]), /No slide titles found/);
});
