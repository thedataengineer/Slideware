const test = require("node:test");
const assert = require("node:assert/strict");
const { searchDeck } = require("../lib-test/features/search.js");

function deck() {
  return {
    slideCount: 2,
    slides: [
      {
        id: "s1",
        index: 1,
        shapes: [
          { id: "a", name: "Title 1", type: "TextBox", left: 0, top: 0, width: 10, height: 10, text: "Revenue grew 40% year over year" },
          { id: "b", name: "Body 1", type: "TextBox", left: 0, top: 0, width: 10, height: 10, text: "" },
        ],
      },
      {
        id: "s2",
        index: 2,
        shapes: [
          { id: "c", name: "Body 2", type: "TextBox", left: 0, top: 0, width: 10, height: 10, text: "Costs fell while REVENUE stabilized across the second half of the fiscal year" },
        ],
      },
    ],
  };
}

test("finds case-insensitive matches across slides in slide order", () => {
  const hits = searchDeck(deck(), "revenue");
  assert.equal(hits.length, 2);
  assert.deepEqual(hits.map((hit) => hit.slideIndex), [1, 2]);
  assert.equal(hits[0].shapeId, "a");
});

test("builds a snippet window around the match", () => {
  const hits = searchDeck(deck(), "stabilized");
  assert.match(hits[0].snippet, /REVENUE stabilized across/);
  assert.ok(hits[0].snippet.length < deck().slides[1].shapes[0].text.length + 10);
});

test("rejects a blank query", () => {
  assert.throws(() => searchDeck(deck(), "   "), /Type something to search/);
});

test("returns no hits when nothing matches", () => {
  assert.deepEqual(searchDeck(deck(), "unicorn"), []);
});
