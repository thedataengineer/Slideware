const test = require("node:test");
const assert = require("node:assert/strict");
const {
  Recorder,
  parseAutomations,
  serializeAutomations,
  validateName,
} = require("../lib-test/features/automations.js");

test("records dispatched steps in order", () => {
  const recorder = new Recorder();
  recorder.start();
  recorder.recordStep("align.left");
  recorder.recordStep("matrix", { columns: 3, horizontalGap: 16, verticalGap: 16 });
  const steps = recorder.stop();
  assert.deepEqual(steps, [
    { op: "align.left", params: undefined },
    { op: "matrix", params: { columns: 3, horizontalGap: 16, verticalGap: 16 } },
  ]);
});

test("ignores steps while not recording", () => {
  const recorder = new Recorder();
  recorder.recordStep("align.left");
  recorder.start();
  const steps = recorder.stop();
  assert.deepEqual(steps, []);
  assert.equal(recorder.isRecording(), false);
});

test("round-trips automations through serialization", () => {
  const automations = [{ name: "Tidy row", steps: [{ op: "align.top" }, { op: "size.width" }] }];
  assert.deepEqual(parseAutomations(serializeAutomations(automations)), automations);
});

test("returns an empty list for corrupt storage", () => {
  assert.deepEqual(parseAutomations("nope"), []);
  assert.deepEqual(parseAutomations(null), []);
  assert.deepEqual(parseAutomations(JSON.stringify({ not: "a list" })), []);
});

test("validates automation names", () => {
  const existing = [{ name: "Tidy row", steps: [] }];
  assert.throws(() => validateName("", existing), /Name the automation/);
  assert.throws(() => validateName("Tidy row", existing), /already exists/);
  assert.equal(validateName("Fresh", existing), "Fresh");
});
