const test = require("node:test");
const assert = require("node:assert/strict");
const {
  MAX_QUESTIONS,
  parseDeckAnalysis,
  standardQuestions,
  withFallbackQuestions,
  defaultAnswers,
  setAnswer,
} = require("../lib-test/features/deck-interview.js");

const GOOD = JSON.stringify({
  summary: "A product briefing.",
  detected: ["5 sections", "Nested bullets"],
  questions: [
    { question: "Who is the audience?", options: ["Internal exec", "Investors"], recommended: "Investors" },
    { question: "How many slides?", options: ["5", "7", "12"], recommended: "7" },
  ],
});

test("parses a well formed analysis and stamps ids and kinds", () => {
  const analysis = parseDeckAnalysis(GOOD);
  assert.equal(analysis.summary, "A product briefing.");
  assert.deepEqual(analysis.detected, ["5 sections", "Nested bullets"]);
  assert.deepEqual(
    analysis.questions.map((q) => q.id),
    ["q1", "q2"]
  );
  analysis.questions.forEach((q) => assert.equal(q.kind, "choice"));
});

test("moves the recommended option to the front so a select defaults to it", () => {
  const analysis = parseDeckAnalysis(GOOD);
  analysis.questions.forEach((q) => assert.equal(q.recommended, q.options[0]));
  assert.equal(analysis.questions[0].options[0], "Investors");
  assert.equal(analysis.questions[1].options[0], "7");
});

test("keeps the first four questions when the model returns too many", () => {
  const many = {
    questions: Array.from({ length: 9 }, (_, i) => ({
      question: `Q${i}`,
      options: ["a", "b"],
      recommended: "a",
    })),
  };
  const analysis = parseDeckAnalysis(JSON.stringify(many));
  assert.equal(analysis.questions.length, MAX_QUESTIONS);
  assert.equal(analysis.questions[0].question, "Q0");
});

test("clamps options to four and keeps the recommendation among them", () => {
  const raw = JSON.stringify({
    questions: [{ question: "Pick", options: ["a", "b", "c", "d", "e", "f", "g"], recommended: "f" }],
  });
  const question = parseDeckAnalysis(raw).questions[0];
  assert.equal(question.options.length, 4);
  assert.equal(question.options[0], "f");
});

test("de-dupes options case insensitively", () => {
  const raw = JSON.stringify({ questions: [{ question: "Pick", options: ["A", "a", "B"], recommended: "B" }] });
  assert.deepEqual(parseDeckAnalysis(raw).questions[0].options, ["B", "A"]);
});

test("adopts a recommendation the model forgot to list", () => {
  const raw = JSON.stringify({
    questions: [{ question: "Pick", options: ["a", "b"], recommended: "Something else" }],
  });
  const question = parseDeckAnalysis(raw).questions[0];
  assert.deepEqual(question.options, ["Something else", "a", "b"]);
  assert.equal(question.recommended, "Something else");
});

test("matches a recommendation loosely despite case and punctuation", () => {
  const raw = JSON.stringify({
    questions: [{ question: "Pick", options: ["Internal exec", "Investors"], recommended: "internal exec." }],
  });
  const question = parseDeckAnalysis(raw).questions[0];
  assert.deepEqual(question.options, ["Internal exec", "Investors"]);
  assert.equal(question.recommended, "Internal exec");
});

test("drops individually bad questions but keeps their siblings", () => {
  const raw = JSON.stringify({
    questions: [
      { question: "", options: ["a", "b"], recommended: "a" },
      { question: "Only one option", options: ["a"], recommended: "a" },
      { question: "No recommendation", options: ["a", "b"], recommended: "" },
      { question: "Good", options: ["a", "b"], recommended: "a" },
    ],
  });
  const analysis = parseDeckAnalysis(raw);
  assert.equal(analysis.questions.length, 1);
  assert.equal(analysis.questions[0].question, "Good");
  assert.equal(analysis.questions[0].id, "q1");
});

test("ignores model supplied ids and kinds", () => {
  const raw = JSON.stringify({
    questions: [{ id: "banana", kind: "multi", question: "Pick", options: ["a", "b"], recommended: "a" }],
  });
  const question = parseDeckAnalysis(raw).questions[0];
  assert.equal(question.id, "q1");
  assert.equal(question.kind, "choice");
});

test("tolerates a missing summary and detected list", () => {
  const analysis = parseDeckAnalysis('{"questions":[{"question":"Q","options":["a","b"],"recommended":"a"}]}');
  assert.equal(analysis.summary, "");
  assert.deepEqual(analysis.detected, []);
});

test("clamps detected observations to five", () => {
  const raw = JSON.stringify({
    detected: ["a", "b", "c", "d", "e", "f", "g", "h"],
    questions: [{ question: "Q", options: ["a", "b"], recommended: "a" }],
  });
  assert.equal(parseDeckAnalysis(raw).detected.length, 5);
});

test("reads a bare top level array as the question list", () => {
  const raw = '[{"question":"Q","options":["a","b"],"recommended":"a"}]';
  assert.equal(parseDeckAnalysis(raw).questions.length, 1);
});

test("throws when the model returned nothing usable", () => {
  assert.throws(() => parseDeckAnalysis("I'd be happy to help!"), /could not be read as JSON/);
  assert.throws(() => parseDeckAnalysis('{"summary":"x"}'), /could not be read as JSON/);
  assert.throws(() => parseDeckAnalysis('{"questions":[{"question":"Who'), /could not be read as JSON/);
});

test("no surviving question is not an error", () => {
  assert.deepEqual(parseDeckAnalysis('{"questions":[]}').questions, []);
  assert.deepEqual(parseDeckAnalysis('{"questions":[{"question":"?"}]}').questions, []);
});

test("the standard question set satisfies the same contract", () => {
  const questions = standardQuestions();
  assert.equal(questions.length, 4);
  const ids = new Set(questions.map((q) => q.id));
  assert.equal(ids.size, questions.length);
  questions.forEach((q) => {
    assert.equal(q.kind, "choice");
    assert.equal(q.recommended, q.options[0]);
    assert.ok(q.options.length >= 2 && q.options.length <= 4, `${q.id} option count`);
  });
});

test("fallback questions fill in only when none survived", () => {
  const empty = withFallbackQuestions({ summary: "", detected: [], questions: [] });
  assert.equal(empty.questions.length, 4);

  const populated = parseDeckAnalysis(GOOD);
  assert.deepEqual(withFallbackQuestions(populated), populated);
});

test("defaultAnswers takes every recommendation and carries the question text", () => {
  const analysis = parseDeckAnalysis(GOOD);
  assert.deepEqual(defaultAnswers(analysis.questions), [
    { questionId: "q1", question: "Who is the audience?", answer: "Investors" },
    { questionId: "q2", question: "How many slides?", answer: "7" },
  ]);
});

test("setAnswer replaces one entry without mutating the original", () => {
  const before = defaultAnswers(parseDeckAnalysis(GOOD).questions);
  const after = setAnswer(before, "q1", "Internal exec");
  assert.equal(after[0].answer, "Internal exec");
  assert.equal(after[1].answer, "7");
  assert.equal(before[0].answer, "Investors");
  assert.notEqual(before, after);
});

test("setAnswer refuses an unknown question", () => {
  const answers = defaultAnswers(parseDeckAnalysis(GOOD).questions);
  assert.throws(() => setAnswer(answers, "q9", "x"), /Unknown question/);
});
