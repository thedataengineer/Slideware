const test = require("node:test");
const assert = require("node:assert/strict");
const { TOOLS, handleMessage } = require("../mcp-server/protocol.js");

test("initialize returns protocol version and server info", () => {
  const { response, effect } = handleMessage({ jsonrpc: "2.0", id: 1, method: "initialize", params: {} });
  assert.equal(effect, undefined);
  assert.equal(response.id, 1);
  assert.equal(response.result.serverInfo.name, "slideware");
  assert.ok(response.result.protocolVersion);
  assert.ok(response.result.capabilities.tools);
});

test("notifications produce no response", () => {
  const { response, effect } = handleMessage({ jsonrpc: "2.0", method: "notifications/initialized" });
  assert.equal(response, undefined);
  assert.equal(effect, undefined);
});

test("tools/list returns the catalog with schemas", () => {
  const { response } = handleMessage({ jsonrpc: "2.0", id: 2, method: "tools/list" });
  assert.equal(response.result.tools.length, TOOLS.length);
  const names = response.result.tools.map((tool) => tool.name);
  assert.ok(names.includes("align_shapes"));
  assert.ok(names.includes("get_deck_outline"));
  response.result.tools.forEach((tool) => {
    assert.equal(typeof tool.description, "string");
    assert.equal(tool.inputSchema.type, "object");
  });
});

test("tools/call defers to an effect for the pane", () => {
  const { response, effect } = handleMessage({
    jsonrpc: "2.0",
    id: 3,
    method: "tools/call",
    params: { name: "align_shapes", arguments: { mode: "left" } },
  });
  assert.equal(response, undefined);
  assert.deepEqual(effect, { type: "forward", id: 3, tool: "align_shapes", args: { mode: "left" } });
});

test("tools/call with an unknown tool errors immediately", () => {
  const { response } = handleMessage({
    jsonrpc: "2.0",
    id: 4,
    method: "tools/call",
    params: { name: "nope", arguments: {} },
  });
  assert.equal(response.error.code, -32602);
});

test("ping returns an empty result and unknown methods error", () => {
  assert.deepEqual(handleMessage({ jsonrpc: "2.0", id: 5, method: "ping" }).response.result, {});
  const { response } = handleMessage({ jsonrpc: "2.0", id: 6, method: "bogus/method" });
  assert.equal(response.error.code, -32601);
});
