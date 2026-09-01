const test = require("node:test");
const assert = require("node:assert/strict");
const { PLACEHOLDER_ADDIN_ID, applyAddinId, assertUuid } = require("../scripts/addin-id.js");

test("swaps every occurrence of the placeholder id", () => {
  const source = JSON.stringify({
    id: PLACEHOLDER_ADDIN_ID,
    note: `see ${PLACEHOLDER_ADDIN_ID} twice`,
  });

  const rewritten = applyAddinId(source, "11111111-2222-4333-8444-555555555555");

  assert.ok(!rewritten.includes(PLACEHOLDER_ADDIN_ID));
  assert.equal((rewritten.match(/11111111-2222-4333-8444-555555555555/g) || []).length, 2);
});

test("leaves a manifest without the placeholder untouched", () => {
  const source = JSON.stringify({ id: "99999999-8888-4777-8666-555555555555" });

  assert.equal(applyAddinId(source, "11111111-2222-4333-8444-555555555555"), source);
});

test("accepts a canonical uuid", () => {
  assert.doesNotThrow(() => assertUuid("fe737f47-102d-4d29-8a47-50844e10ac76"));
});

test("rejects anything that is not a uuid", () => {
  ["", "nope", "fe737f47102d4d298a4750844e10ac76", "fe737f47-102d-4d29-8a47", "zz737f47-102d-4d29-8a47-50844e10ac76"].forEach(
    (value) => {
      assert.throws(() => assertUuid(value), /add-in id/i, `expected ${value} to be rejected`);
    }
  );
});
