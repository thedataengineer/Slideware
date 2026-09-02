const test = require("node:test");
const assert = require("node:assert/strict");
const { extractJson, cleanString, cleanStringList } = require("../lib-test/features/model-json.js");

test("parses a bare object and a bare array", () => {
  assert.deepEqual(extractJson('{"a":1}'), { a: 1 });
  assert.deepEqual(extractJson('[1,2]'), [1, 2]);
});

test("parses a fenced block", () => {
  assert.deepEqual(extractJson('```json\n{"a":1}\n```'), { a: 1 });
  assert.deepEqual(extractJson('```\n{"a":1}\n```'), { a: 1 });
});

test("parses a fence buried in chat preamble and trailing prose", () => {
  const raw = 'Sure! Here you go:\n```json\n{"a":1}\n```\nHope that helps!';
  assert.deepEqual(extractJson(raw), { a: 1 });
});

test("parses unfenced json followed by prose", () => {
  assert.deepEqual(extractJson('{"a":1}\n\nLet me know if you want changes.'), { a: 1 });
});

test("repairs a single trailing comma", () => {
  assert.deepEqual(extractJson('{"a":[1,2,],}'), { a: [1, 2] });
});

test("a brace inside a string does not end the slice", () => {
  assert.deepEqual(extractJson('{"a":"x } y"}'), { a: "x } y" });
});

test("an escaped quote inside a string does not end the string", () => {
  assert.deepEqual(extractJson('{"a":"say \\"hi\\""}'), { a: 'say "hi"' });
});

test("a comma inside a string is never treated as a trailing comma", () => {
  assert.deepEqual(extractJson('{"a":"one, two"}'), { a: "one, two" });
});

test("returns null instead of throwing on unusable input", () => {
  assert.equal(extractJson("not json"), null);
  assert.equal(extractJson(""), null);
  assert.equal(extractJson('{"a": '), null);
  assert.equal(extractJson(42), null);
});

test("cleanString trims, collapses whitespace, and clamps", () => {
  assert.equal(cleanString("  a  b  ", 10), "a b");
  assert.equal(cleanString("a\n\nb", 10), "a b");
  assert.equal(cleanString("abcdefghij", 5), "abcde");
  assert.equal(cleanString(42, 10), "");
  assert.equal(cleanString(undefined, 10), "");
});

test("cleanStringList drops empties and non-strings, de-dupes case-insensitively", () => {
  assert.deepEqual(cleanStringList(["a", "A", "", "b", 7], 10, 50), ["a", "b"]);
});

test("cleanStringList clamps the item count", () => {
  assert.deepEqual(cleanStringList(["a", "b", "c"], 2, 50), ["a", "b"]);
});

test("cleanStringList returns empty for a non-array", () => {
  assert.deepEqual(cleanStringList("nope", 5, 50), []);
});
