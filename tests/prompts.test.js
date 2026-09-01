const test = require("node:test");
const assert = require("node:assert/strict");
const {
  presetPrompt,
  editPrompt,
  createPrompt,
  parseCreateResponse,
  translatePrompt,
  darwinSystem,
} = require("../lib-test/features/prompts.js");

test("preset prompts vary by mode and carry the slide text", () => {
  const proofread = presetPrompt("Teh quick fox", "proofread");
  assert.match(proofread.user, /Teh quick fox/);
  assert.match(proofread.user, /typos/i);
  const shorten = presetPrompt("Some very long text", "shorten");
  assert.match(shorten.user, /shorter/i);
  assert.notEqual(proofread.user, shorten.user);
  assert.throws(() => presetPrompt("x", "nope"), /Unknown preset/);
});

test("freeform edit prompt embeds the instruction", () => {
  const prompt = editPrompt("Original", "Make it confident");
  assert.match(prompt.user, /Original/);
  assert.match(prompt.user, /Make it confident/);
});

test("edit prompt carries deck context when provided", () => {
  const prompt = editPrompt("Original", "Align with the deck", "Slide 1: Q3 Results\nRevenue grew");
  assert.match(prompt.system, /Slide 1: Q3 Results/);
  const bare = editPrompt("Original", "Align with the deck");
  assert.doesNotMatch(bare.system, /deck outline/i);
});

test("create prompt grounds content in the deck outline and truncates it", () => {
  const prompt = createPrompt("agenda", "Slide 2: Hiring Plan\nWe need 4 engineers");
  assert.match(prompt.system, /Slide 2: Hiring Plan/);
  assert.match(prompt.system, /consistent with/i);

  const long = createPrompt("agenda", "x".repeat(20000));
  assert.ok(long.system.length < 10000);
  assert.match(long.system, /outline truncated/);
});

test("create prompt demands strict JSON and parser accepts fenced output", () => {
  const prompt = createPrompt("Q3 results");
  assert.match(prompt.user, /Q3 results/);
  assert.match(prompt.system, /JSON/);

  const bare = parseCreateResponse('{"title": "T", "bullets": ["a", "b"]}');
  assert.deepEqual(bare, { title: "T", bullets: ["a", "b"] });

  const fenced = parseCreateResponse('```json\n{"title": "T", "bullets": ["a"]}\n```');
  assert.deepEqual(fenced, { title: "T", bullets: ["a"] });

  assert.throws(() => parseCreateResponse("not json"), /could not parse/i);
  assert.throws(() => parseCreateResponse('{"title": 3}'), /could not parse/i);
});

test("translate prompt names the target language", () => {
  const prompt = translatePrompt("Hello", "Japanese");
  assert.match(prompt.user, /Japanese/);
  assert.match(prompt.user, /Hello/);
});

test("darwin system prompt truncates long outlines", () => {
  const outline = "x".repeat(10000);
  const system = darwinSystem(outline);
  assert.ok(system.length < 9000);
  assert.match(system, /presentation/i);
});
