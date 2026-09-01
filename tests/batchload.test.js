const test = require("node:test");
const assert = require("node:assert/strict");
const { loadInBatches } = require("../lib-test/features/batchload.js");

function fakeHost(badItems) {
  const bad = new Set(badItems);
  const state = { queued: [], syncs: 0, loaded: new Set() };
  return {
    state,
    queue: (item) => state.queued.push(item),
    sync: async () => {
      state.syncs += 1;
      const batch = state.queued;
      state.queued = [];
      const offender = batch.find((item) => bad.has(item));
      if (offender) throw new Error(`InvalidArgument at ${offender}`);
      batch.forEach((item) => state.loaded.add(item));
    },
  };
}

test("returns nothing and skips the host when there is no work", async () => {
  const host = fakeHost([]);
  const failed = await loadInBatches({ items: [], queue: host.queue, sync: host.sync });

  assert.deepEqual(failed, []);
  assert.equal(host.state.syncs, 0);
});

test("loads a clean batch in a single sync", async () => {
  const items = ["a", "b", "c", "d"];
  const host = fakeHost([]);
  const failed = await loadInBatches({ items, queue: host.queue, sync: host.sync });

  assert.deepEqual(failed, []);
  assert.equal(host.state.syncs, 1);
  assert.deepEqual([...host.state.loaded].sort(), items);
});

test("isolates the offending item and still loads every other shape", async () => {
  const items = ["a", "b", "bad", "d", "e", "f", "g", "h"];
  const host = fakeHost(["bad"]);
  const failed = await loadInBatches({ items, queue: host.queue, sync: host.sync });

  assert.deepEqual(failed, ["bad"]);
  assert.deepEqual([...host.state.loaded].sort(), ["a", "b", "d", "e", "f", "g", "h"]);
});

test("bisects instead of probing every item one at a time", async () => {
  const items = Array.from({ length: 64 }, (_, index) => `shape-${index}`);
  const host = fakeHost(["shape-40"]);
  const failed = await loadInBatches({ items, queue: host.queue, sync: host.sync });

  assert.deepEqual(failed, ["shape-40"]);
  assert.ok(host.state.syncs < items.length, `expected fewer than ${items.length} syncs`);
});

test("reports every item when the whole batch is unsupported", async () => {
  const items = ["x", "y", "z"];
  const host = fakeHost(items);
  const failed = await loadInBatches({ items, queue: host.queue, sync: host.sync });

  assert.deepEqual(failed.sort(), items);
  assert.equal(host.state.loaded.size, 0);
});

test("keeps failures in the original item order", async () => {
  const items = ["a", "bad1", "c", "bad2", "e"];
  const host = fakeHost(["bad1", "bad2"]);
  const failed = await loadInBatches({ items, queue: host.queue, sync: host.sync });

  assert.deepEqual(failed, ["bad1", "bad2"]);
  assert.deepEqual([...host.state.loaded].sort(), ["a", "c", "e"]);
});
