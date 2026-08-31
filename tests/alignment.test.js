const test = require("node:test");
const assert = require("node:assert/strict");
const { alignShapes, distributeShapes, arrangeMatrix, arrangeCircle } = require("../lib-test/alignment.js");

function loadPowerPointAdapter({ shapes, supported = true, hostError }) {
  const requirementCalls = [];
  const runCalls = [];
  global.Office = {
    context: {
      requirements: {
        isSetSupported: (...args) => {
          requirementCalls.push(args);
          return supported;
        },
      },
    },
  };
  global.PowerPoint = {
    run: async (callback) => {
      runCalls.push(callback);
      if (hostError) throw hostError;
      return callback({
        presentation: {
          getSelectedShapes: () => ({
            items: shapes,
            load: () => undefined,
          }),
        },
        sync: async () => undefined,
      });
    },
  };
  delete require.cache[require.resolve("../lib-test/powerpoint.js")];
  return { ...require("../lib-test/powerpoint.js"), requirementCalls, runCalls };
}

const shapes = [
  { id: "a", left: 10, top: 20, width: 20, height: 10 },
  { id: "b", left: 50, top: 60, width: 40, height: 30 },
];

test("aligns mixed-size shapes to each selection edge", () => {
  assert.deepEqual(alignShapes(shapes, "left"), [
    { id: "a", left: 10, top: 20 },
    { id: "b", left: 10, top: 60 },
  ]);
  assert.deepEqual(alignShapes(shapes, "right"), [
    { id: "a", left: 70, top: 20 },
    { id: "b", left: 50, top: 60 },
  ]);
  assert.deepEqual(alignShapes(shapes, "top"), [
    { id: "a", left: 10, top: 20 },
    { id: "b", left: 50, top: 20 },
  ]);
  assert.deepEqual(alignShapes(shapes, "bottom"), [
    { id: "a", left: 10, top: 80 },
    { id: "b", left: 50, top: 60 },
  ]);
});

test("aligns shape centers to selection centerlines", () => {
  assert.deepEqual(alignShapes(shapes, "center"), [
    { id: "a", left: 40, top: 20 },
    { id: "b", left: 30, top: 60 },
  ]);
  assert.deepEqual(alignShapes(shapes, "middle"), [
    { id: "a", left: 10, top: 50 },
    { id: "b", left: 50, top: 40 },
  ]);
});

test("rejects alignment with fewer than two shapes", () => {
  assert.throws(() => alignShapes(shapes.slice(0, 1), "left"), /Select at least 2 shapes/);
});

test("aligns shapes when selection uses negative coordinates", () => {
  const input = [
    { id: "a", left: -30, top: -10, width: 10, height: 10 },
    { id: "b", left: 20, top: 10, width: 20, height: 20 },
  ];
  assert.deepEqual(alignShapes(input, "left"), [
    { id: "a", left: -30, top: -10 },
    { id: "b", left: -30, top: 10 },
  ]);
});

test("distributes horizontal gaps while preserving outer shape positions", () => {
  const input = [
    { id: "right", left: 100, top: 0, width: 20, height: 10 },
    { id: "left", left: 0, top: 0, width: 10, height: 10 },
    { id: "middle", left: 35, top: 0, width: 30, height: 10 },
  ];
  assert.deepEqual(distributeShapes(input, "horizontal"), [
    { id: "right", left: 100, top: 0 },
    { id: "left", left: 0, top: 0 },
    { id: "middle", left: 40, top: 0 },
  ]);
});

test("distributes vertical gaps with mixed heights", () => {
  const input = [
    { id: "top", left: 0, top: 0, width: 10, height: 20 },
    { id: "bottom", left: 0, top: 100, width: 10, height: 10 },
    { id: "middle", left: 0, top: 30, width: 10, height: 30 },
  ];
  assert.deepEqual(distributeShapes(input, "vertical"), [
    { id: "top", left: 0, top: 0 },
    { id: "bottom", left: 0, top: 100 },
    { id: "middle", left: 0, top: 45 },
  ]);
});

test("uses equal negative gaps when shapes exceed available span", () => {
  const input = [
    { id: "a", left: 0, top: 0, width: 50, height: 10 },
    { id: "b", left: 30, top: 0, width: 50, height: 10 },
    { id: "c", left: 60, top: 0, width: 50, height: 10 },
  ];
  assert.equal(distributeShapes(input, "horizontal").find((shape) => shape.id === "b").left, 30);
});

test("rejects distribution with fewer than three shapes", () => {
  assert.throws(() => distributeShapes(shapes, "horizontal"), /Select at least 3 shapes/);
});

test("arranges stable reading order into mixed-size matrix cells", () => {
  const input = [
    { id: "c", left: 0, top: 50, width: 20, height: 10 },
    { id: "a", left: 0, top: 0, width: 10, height: 10 },
    { id: "b", left: 50, top: 0, width: 30, height: 20 },
  ];
  assert.deepEqual(arrangeMatrix(input, { columns: 2, horizontalGap: 10, verticalGap: 5 }), [
    { id: "c", left: 0, top: 25 },
    { id: "a", left: 5, top: 5 },
    { id: "b", left: 30, top: 0 },
  ]);
});

test("rejects invalid matrix settings", () => {
  assert.throws(
    () => arrangeMatrix(shapes, { columns: 3, horizontalGap: 10, verticalGap: 10 }),
    /Columns cannot exceed selected shape count/,
  );
  assert.throws(
    () => arrangeMatrix(shapes, { columns: 2.5, horizontalGap: 10, verticalGap: 10 }),
    /Columns must be a positive integer/,
  );
});

test("arranges shape centers clockwise using radius and start angle", () => {
  const input = [
    { id: "a", left: 0, top: 0, width: 10, height: 10 },
    { id: "b", left: 20, top: 0, width: 10, height: 10 },
    { id: "c", left: 40, top: 0, width: 10, height: 10 },
    { id: "d", left: 60, top: 0, width: 10, height: 10 },
  ];
  assert.deepEqual(arrangeCircle(input, { radius: 20, startAngle: -90, clockwise: true }), [
    { id: "a", left: 30, top: -20 },
    { id: "b", left: 50, top: 0 },
    { id: "c", left: 30, top: 20 },
    { id: "d", left: 10, top: 0 },
  ]);
});

test("reverses circle direction", () => {
  const input = [
    { id: "a", left: 0, top: 0, width: 10, height: 10 },
    { id: "b", left: 20, top: 0, width: 10, height: 10 },
    { id: "c", left: 40, top: 0, width: 10, height: 10 },
    { id: "d", left: 60, top: 0, width: 10, height: 10 },
  ];
  assert.equal(
    arrangeCircle(input, { radius: 20, startAngle: -90, clockwise: false })[1].left,
    10,
  );
});

test("rejects non-positive circle radius", () => {
  assert.throws(
    () => arrangeCircle(shapes, { radius: 0, startAngle: 0, clockwise: true }),
    /Radius must be greater than 0/,
  );
});

test("applies a computed position to every selected PowerPoint shape", async () => {
  const shapes = [
    { id: "a", left: 10, top: 20, width: 20, height: 10, load: () => undefined },
    { id: "b", left: 50, top: 60, width: 40, height: 30, load: () => undefined },
  ];
  const { applyLayout, requirementCalls } = loadPowerPointAdapter({ shapes });

  const updated = await applyLayout((bounds) =>
    bounds.map((shape) => ({ id: shape.id, left: shape.left + 5, top: shape.top - 5 })),
  );

  assert.equal(updated, 2);
  assert.deepEqual(requirementCalls, [["PowerPointApi", "1.5"]]);
  assert.deepEqual(
    shapes.map(({ id, left, top }) => ({ id, left, top })),
    [
      { id: "a", left: 15, top: 15 },
      { id: "b", left: 55, top: 55 },
    ],
  );
});

test("rejects unsupported PowerPoint hosts before loading the selection", async () => {
  const { applyLayout, SmartAlignmentError, requirementCalls, runCalls } = loadPowerPointAdapter({
    shapes: [],
    supported: false,
  });

  await assert.rejects(
    applyLayout(() => []),
    (error) =>
      error instanceof SmartAlignmentError && error.message === "Smart Alignment requires PowerPoint API 1.5.",
  );
  assert.deepEqual(requirementCalls, [["PowerPointApi", "1.5"]]);
  assert.equal(runCalls.length, 0);
});

test("translates PowerPoint host failures to a stable alignment error", async () => {
  const { applyLayout, SmartAlignmentError } = loadPowerPointAdapter({
    shapes: [],
    hostError: new Error("OfficeExtension.Error: GeneralException"),
  });

  await assert.rejects(
    applyLayout(() => []),
    (error) => error instanceof SmartAlignmentError && error.message === "PowerPoint could not update selected shapes.",
  );
});

test("preserves geometry validation errors from the layout computer", async () => {
  const shapes = [{ id: "a", left: 10, top: 20, width: 20, height: 10, load: () => undefined }];
  const { applyLayout, SmartAlignmentError } = loadPowerPointAdapter({ shapes });

  await assert.rejects(
    applyLayout((bounds) => alignShapes(bounds, "left")),
    (error) => error instanceof Error && !(error instanceof SmartAlignmentError) && error.message === "Select at least 2 shapes.",
  );
});

test("rejects a missing shape position before queuing any coordinate changes", async () => {
  const shapes = [
    { id: "a", left: 10, top: 20, width: 20, height: 10, load: () => undefined },
    { id: "b", left: 50, top: 60, width: 40, height: 30, load: () => undefined },
  ];
  const { applyLayout, SmartAlignmentError } = loadPowerPointAdapter({ shapes });

  await assert.rejects(
    applyLayout(() => [{ id: "a", left: 15, top: 15 }]),
    (error) => error instanceof SmartAlignmentError && error.message === "Layout returned no position for shape b.",
  );
  assert.deepEqual(
    shapes.map(({ id, left, top }) => ({ id, left, top })),
    [
      { id: "a", left: 10, top: 20 },
      { id: "b", left: 50, top: 60 },
    ],
  );
});
