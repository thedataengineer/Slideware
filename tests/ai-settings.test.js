const test = require("node:test");
const assert = require("node:assert/strict");
const {
  defaultAiSettings,
  parseAiSettings,
  serializeAiSettings,
} = require("../lib-test/features/ai-settings.js");

test("round-trips settings through serialization", () => {
  const settings = {
    provider: "ollama",
    apiKey: "sk-ant-x",
    ollamaUrl: "http://localhost:11434",
    ollamaModel: "llama3.2",
  };
  assert.deepEqual(parseAiSettings(serializeAiSettings(settings)), settings);
});

test("falls back to defaults on corrupt or missing input", () => {
  assert.deepEqual(parseAiSettings(null), defaultAiSettings());
  assert.deepEqual(parseAiSettings("not json"), defaultAiSettings());
  assert.deepEqual(parseAiSettings(JSON.stringify({ provider: 5 })), defaultAiSettings());
});

test("rejects unknown providers but keeps valid fields", () => {
  const parsed = parseAiSettings(JSON.stringify({ provider: "openai", apiKey: "k" }));
  assert.equal(parsed.provider, "claude");
  assert.equal(parsed.apiKey, "k");
});

test("strips a trailing slash from the ollama url", () => {
  const parsed = parseAiSettings(
    JSON.stringify({ provider: "ollama", ollamaUrl: "http://127.0.0.1:11434/" })
  );
  assert.equal(parsed.ollamaUrl, "http://127.0.0.1:11434");
});

test("defaults choose claude with the standard ollama endpoint", () => {
  const defaults = defaultAiSettings();
  assert.equal(defaults.provider, "claude");
  assert.equal(defaults.ollamaUrl, "http://localhost:11434");
  assert.equal(defaults.ollamaModel, "llama3.2");
  assert.equal(defaults.apiKey, "");
});
