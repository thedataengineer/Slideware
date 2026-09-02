const test = require("node:test");
const assert = require("node:assert/strict");
const {
  MAX_SLIDES,
  MAX_BULLETS,
  parseSlidePlan,
  movePlanSlide,
  removePlanSlide,
  updatePlanSlide,
  planSummary,
  slideMeta,
} = require("../lib-test/features/deck-plan.js");

const GOOD = JSON.stringify({
  slides: [
    { kind: "title", title: "Money Moxie", bullets: [] },
    { kind: "section", title: "Architecture", bullets: [] },
    { kind: "bullets", title: "Stack", bullets: ["Railway", "Plaid"], notes: "Keep it short." },
  ],
});

function plan() {
  return parseSlidePlan(GOOD);
}

test("parses a well formed plan", () => {
  assert.deepEqual(plan().slides, [
    { kind: "title", title: "Money Moxie", bullets: [] },
    { kind: "section", title: "Architecture", bullets: [] },
    { kind: "bullets", title: "Stack", bullets: ["Railway", "Plaid"], notes: "Keep it short." },
  ]);
});

test("omits notes entirely when the model sent an empty string", () => {
  const parsed = parseSlidePlan('{"slides":[{"kind":"bullets","title":"T","bullets":["a"],"notes":"  "}]}');
  assert.equal("notes" in parsed.slides[0], false);
});

test("maps kind synonyms onto the three supported kinds", () => {
  const raw = JSON.stringify({
    slides: [
      { kind: "cover", title: "A", bullets: [] },
      { kind: "divider", title: "B", bullets: [] },
      { kind: "content", title: "C", bullets: ["x"] },
      { kind: "TITLE ", title: "D", bullets: [] },
    ],
  });
  assert.deepEqual(
    parseSlidePlan(raw).slides.map((s) => s.kind),
    ["title", "section", "bullets", "title"]
  );
});

test("infers a kind the model omitted or garbled", () => {
  const raw = JSON.stringify({
    slides: [
      { kind: "nonsense", title: "Has bullets", bullets: ["x"] },
      { title: "No bullets", bullets: [] },
    ],
  });
  assert.deepEqual(
    parseSlidePlan(raw).slides.map((s) => s.kind),
    ["bullets", "section"]
  );
});

test("demotes an empty bullets slide to a section rather than dropping it", () => {
  const parsed = parseSlidePlan('{"slides":[{"kind":"bullets","title":"Kept","bullets":[]}]}');
  assert.equal(parsed.slides.length, 1);
  assert.equal(parsed.slides[0].kind, "section");
  assert.equal(parsed.slides[0].title, "Kept");
});

test("strips bullet glyphs the model added", () => {
  const raw = JSON.stringify({
    slides: [{ kind: "bullets", title: "T", bullets: ["- One", "* Two", "• Three", "1. Four", "2) Five"] }],
  });
  assert.deepEqual(parseSlidePlan(raw).slides[0].bullets, ["One", "Two", "Three", "Four", "Five"]);
});

test("clamps bullets per slide and slides per plan", () => {
  const many = JSON.stringify({
    slides: [{ kind: "bullets", title: "T", bullets: Array.from({ length: 9 }, (_, i) => `b${i}`) }],
  });
  assert.equal(parseSlidePlan(many).slides[0].bullets.length, MAX_BULLETS);

  const lots = JSON.stringify({
    slides: Array.from({ length: 50 }, (_, i) => ({ kind: "section", title: `S${i}`, bullets: [] })),
  });
  assert.equal(parseSlidePlan(lots).slides.length, MAX_SLIDES);
});

test("drops a title-less slide but keeps its siblings", () => {
  const raw = JSON.stringify({
    slides: [
      { kind: "section", title: "   ", bullets: [] },
      { kind: "section", title: "Kept", bullets: [] },
    ],
  });
  const parsed = parseSlidePlan(raw);
  assert.equal(parsed.slides.length, 1);
  assert.equal(parsed.slides[0].title, "Kept");
});

test("accepts the plan under several container keys", () => {
  assert.equal(parseSlidePlan('{"plan":[{"kind":"section","title":"A"}]}').slides.length, 1);
  assert.equal(parseSlidePlan('[{"kind":"section","title":"A"}]').slides.length, 1);
});

test("throws when nothing usable came back", () => {
  assert.throws(() => parseSlidePlan("here is your deck!"), /could not be read as JSON/);
  assert.throws(() => parseSlidePlan('{"slides":[]}'), /no usable slides/);
  assert.throws(() => parseSlidePlan('{"slides":[{"bullets":["a"]}]}'), /no usable slides/);
});

test("movePlanSlide swaps neighbours without mutating the original", () => {
  const before = plan();
  const snapshot = JSON.parse(JSON.stringify(before));
  const after = movePlanSlide(before, 1, -1);

  assert.deepEqual(
    after.slides.map((s) => s.title),
    ["Architecture", "Money Moxie", "Stack"]
  );
  assert.deepEqual(before, snapshot);
});

test("moving past either end is a no-op, not an error", () => {
  const before = plan();
  assert.deepEqual(movePlanSlide(before, 0, -1), before);
  assert.deepEqual(movePlanSlide(before, 2, 1), before);
});

test("edit operations reject an out of range index", () => {
  assert.throws(() => movePlanSlide(plan(), 99, -1), /slide/i);
  assert.throws(() => removePlanSlide(plan(), 99), /slide/i);
  assert.throws(() => updatePlanSlide(plan(), 99, { title: "x" }), /slide/i);
});

test("removePlanSlide drops exactly one and leaves the original alone", () => {
  const before = plan();
  const after = removePlanSlide(before, 1);
  assert.deepEqual(
    after.slides.map((s) => s.title),
    ["Money Moxie", "Stack"]
  );
  assert.equal(before.slides.length, 3);
});

test("updatePlanSlide applies the same clamps to hand typed content", () => {
  const after = updatePlanSlide(plan(), 2, {
    title: "  Trimmed  ",
    bullets: ["- glyph", "", "x".repeat(400), "a", "b", "c", "d", "e"],
  });
  const slide = after.slides[2];
  assert.equal(slide.title, "Trimmed");
  assert.equal(slide.bullets.length, MAX_BULLETS);
  assert.equal(slide.bullets[0], "glyph");
  assert.ok(slide.bullets[1].length <= 200);
});

test("summaries describe the plan and each row", () => {
  assert.equal(planSummary(plan()), "3 slides · 1 title, 1 section, 1 content");
  assert.equal(slideMeta(plan().slides[0]), "Title slide");
  assert.equal(slideMeta(plan().slides[1]), "Section divider");
  assert.equal(slideMeta(plan().slides[2]), "Bullets · 2 bullets · has notes");
});

const { validatePlan, bodyTextFor, summarizeResult } = require("../lib-test/features/deck-plan.js");

test("validatePlan drops unusable slides and refuses an empty plan", () => {
  const cleaned = validatePlan([
    { kind: "bullets", title: "  Kept  ", bullets: ["a"] },
    { kind: "bullets", title: "   ", bullets: ["b"] },
  ]);
  assert.equal(cleaned.length, 1);
  assert.equal(cleaned[0].title, "Kept");

  assert.throws(() => validatePlan([]), /no slides/i);
  assert.throws(() => validatePlan(Array.from({ length: 41 }, () => ({ kind: "section", title: "x", bullets: [] }))), /up to 40/);
});

test("bodyTextFor never adds its own bullet glyphs", () => {
  const body = bodyTextFor({ kind: "bullets", title: "T", bullets: ["one", "two"] });
  assert.equal(body.text, "one\ntwo");
  assert.equal(body.useBullets, true);
  assert.doesNotMatch(body.text, /•/);
});

test("bodyTextFor leaves a section empty and a title unbulleted", () => {
  assert.deepEqual(bodyTextFor({ kind: "section", title: "T", bullets: ["ignored"] }), {
    text: "",
    useBullets: false,
  });
  assert.deepEqual(bodyTextFor({ kind: "title", title: "T", bullets: ["A subtitle"] }), {
    text: "A subtitle",
    useBullets: false,
  });
});

test("summarizeResult reports success, failures, and blanks", () => {
  assert.match(summarizeResult({ added: 6, failed: [], blankSlidesLeft: 0, notesSkipped: 0, fallbackTextBoxes: 0 }), /Added 6 slides/);

  const mixed = summarizeResult({
    added: 4,
    failed: [{ index: 3, title: "Stack", reason: "PowerPoint refused the placeholder text" }],
    blankSlidesLeft: 1,
    notesSkipped: 2,
    fallbackTextBoxes: 0,
  });
  assert.match(mixed, /Added 4 slides/);
  assert.match(mixed, /slide 3/);
  assert.match(mixed, /1 blank slide/);
  assert.match(mixed, /2 notes/);
});
