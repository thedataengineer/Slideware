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
    claudeModel: "claude-sonnet-5",
    ollamaUrl: "http://localhost:11434",
    ollamaModel: "llama3.2",
    openaiUrl: "https://api.groq.com/openai/v1",
    openaiKey: "gsk-x",
    openaiModel: "llama-3.3-70b-versatile",
  };
  assert.deepEqual(parseAiSettings(serializeAiSettings(settings)), settings);
});

test("accepts the openai provider and normalizes its url", () => {
  const parsed = parseAiSettings(
    JSON.stringify({ provider: "openai", openaiUrl: "https://api.openai.com/v1/" })
  );
  assert.equal(parsed.provider, "openai");
  assert.equal(parsed.openaiUrl, "https://api.openai.com/v1");
  assert.equal(parsed.openaiModel, "gpt-4o-mini");
});

test("defaults the claude model and keeps custom ones", () => {
  assert.equal(parseAiSettings(null).claudeModel, "claude-opus-5");
  assert.equal(
    parseAiSettings(JSON.stringify({ claudeModel: "claude-fable-5" })).claudeModel,
    "claude-fable-5"
  );
  assert.equal(parseAiSettings(JSON.stringify({ claudeModel: "  " })).claudeModel, "claude-opus-5");
});

test("falls back to defaults on corrupt or missing input", () => {
  assert.deepEqual(parseAiSettings(null), defaultAiSettings());
  assert.deepEqual(parseAiSettings("not json"), defaultAiSettings());
  assert.deepEqual(parseAiSettings(JSON.stringify({ provider: 5 })), defaultAiSettings());
});

test("rejects unknown providers but keeps valid fields", () => {
  const parsed = parseAiSettings(JSON.stringify({ provider: "bedrock", apiKey: "k" }));
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
  assert.equal(defaults.ollamaUrl, "/ollama");
  assert.equal(defaults.ollamaModel, "llama3.2");
  assert.equal(defaults.apiKey, "");
});
