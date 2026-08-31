# Smart Alignment Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add working alignment, distribution, matrix, and circle layouts for selected PowerPoint shapes.

**Architecture:** Keep coordinate calculations in a pure TypeScript module. Use a thin Office.js adapter to load selected shapes and write positions in one batch, while task-pane code owns controls, validation feedback, and busy state.

**Tech Stack:** TypeScript 5.4, Office.js PowerPoint API 1.5, HTML, CSS, Node built-in test runner, Webpack 5

**Spec:** `docs/superpowers/specs/2026-08-31-smart-alignment-design.md`

## Global Constraints

- Selection access requires PowerPoint API 1.5. Shape position reads and writes require API 1.4.
- Shape width, height, rotation, styling, and grouping must remain unchanged.
- Matrix and circle order shapes by `top`, then `left`, with stable input order for equal coordinates.
- Validate all settings before writing any shape coordinates.
- No backend, authentication, telemetry, remote service, or new runtime dependency.
- Current workspace has no `.git` directory. Commit steps apply only after repository metadata exists.

## File Map

- Create `src/taskpane/alignment.ts`: shape models, validation, geometry calculations.
- Modify `src/taskpane/powerpoint.ts`: Office.js selection adapter and coordinate writer.
- Modify `src/taskpane/taskpane.ts`: DOM bindings, operation dispatch, busy state, status output.
- Modify `src/taskpane/taskpane.html`: Smart Alignment controls.
- Modify `src/taskpane/taskpane.css`: task-pane layout and interaction states.
- Modify `webpack.config.js`: emit task-pane CSS and remove unused command bundles.
- Modify `manifest.json`: replace starter branding and expose Smart Alignment pane only.
- Create `tests/alignment.test.js`: pure geometry regression tests.
- Create `tsconfig.test.json`: compile geometry module for Node tests.
- Modify `package.json`: add `test` and `typecheck` scripts.

---

### Task 1: Alignment Geometry and Test Harness

**Files:**
- Create: `src/taskpane/alignment.ts`
- Create: `tests/alignment.test.js`
- Create: `tsconfig.test.json`
- Modify: `package.json`

**Interfaces:**
- Produces: `ShapeBounds`, `ShapePosition`, `AlignMode`, `alignShapes(shapes, mode)`.
- Consumes: no earlier task interfaces.

- [x] **Step 1: Add test scripts and isolated test compilation**

Add these scripts to `package.json`:

```json
"test": "tsc -p tsconfig.test.json && node --test tests/alignment.test.js",
"typecheck": "tsc --noEmit"
```

Create `tsconfig.test.json`:

```json
{
  "compilerOptions": {
    "module": "commonjs",
    "target": "es2019",
    "outDir": "lib-test",
    "strict": true,
    "esModuleInterop": true
  },
  "include": ["src/taskpane/alignment.ts"]
}
```

- [x] **Step 2: Write failing alignment tests**

Create `tests/alignment.test.js`:

```js
const test = require("node:test");
const assert = require("node:assert/strict");
const { alignShapes } = require("../lib-test/alignment.js");

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
    { id: "a", left: 10, top: 55 },
    { id: "b", left: 50, top: 45 },
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
```

- [x] **Step 3: Run tests and confirm missing module failure**

Run: `npm test`

Expected: FAIL because `src/taskpane/alignment.ts` or exported `alignShapes` does not exist.

- [x] **Step 4: Implement shape models, bounds, and six alignments**

Create `src/taskpane/alignment.ts`:

```ts
export interface ShapeBounds {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ShapePosition {
  id: string;
  left: number;
  top: number;
}

export type AlignMode = "left" | "center" | "right" | "top" | "middle" | "bottom";

function requireShapes(shapes: ShapeBounds[], minimum: number): void {
  if (shapes.length < minimum) {
    throw new Error(`Select at least ${minimum} shapes.`);
  }
}

function selectionBounds(shapes: ShapeBounds[]) {
  return {
    left: Math.min(...shapes.map((shape) => shape.left)),
    top: Math.min(...shapes.map((shape) => shape.top)),
    right: Math.max(...shapes.map((shape) => shape.left + shape.width)),
    bottom: Math.max(...shapes.map((shape) => shape.top + shape.height)),
  };
}

export function alignShapes(shapes: ShapeBounds[], mode: AlignMode): ShapePosition[] {
  requireShapes(shapes, 2);
  const bounds = selectionBounds(shapes);
  const centerX = (bounds.left + bounds.right) / 2;
  const centerY = (bounds.top + bounds.bottom) / 2;

  return shapes.map((shape) => {
    let left = shape.left;
    let top = shape.top;

    if (mode === "left") left = bounds.left;
    if (mode === "center") left = centerX - shape.width / 2;
    if (mode === "right") left = bounds.right - shape.width;
    if (mode === "top") top = bounds.top;
    if (mode === "middle") top = centerY - shape.height / 2;
    if (mode === "bottom") top = bounds.bottom - shape.height;

    return { id: shape.id, left, top };
  });
}
```

- [x] **Step 5: Run geometry tests and typecheck**

Run: `npm test`

Expected: 4 tests PASS.

Run: `npm run typecheck`

Expected: exit code 0.

- [x] **Step 6: Commit when Git metadata exists**

```bash
git add package.json tsconfig.test.json src/taskpane/alignment.ts tests/alignment.test.js
git commit -m "feat: add shape alignment geometry"
```

---

### Task 2: Equal-Gap Distribution

**Files:**
- Modify: `src/taskpane/alignment.ts`
- Modify: `tests/alignment.test.js`

**Interfaces:**
- Consumes: `ShapeBounds`, `ShapePosition`, and `requireShapes` from Task 1.
- Produces: `DistributionAxis`, `distributeShapes(shapes, axis)`.

- [x] **Step 1: Add failing horizontal and vertical distribution tests**

Append to `tests/alignment.test.js` and add `distributeShapes` to its import:

```js
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
```

- [x] **Step 2: Run tests and confirm missing export failure**

Run: `npm test`

Expected: FAIL because `distributeShapes` is undefined.

- [x] **Step 3: Implement equal edge-to-edge distribution**

Append to `src/taskpane/alignment.ts`:

```ts
export type DistributionAxis = "horizontal" | "vertical";

export function distributeShapes(shapes: ShapeBounds[], axis: DistributionAxis): ShapePosition[] {
  requireShapes(shapes, 3);
  const horizontal = axis === "horizontal";
  const sorted = shapes
    .map((shape, index) => ({ shape, index }))
    .sort((a, b) => {
      const aCenter = horizontal
        ? a.shape.left + a.shape.width / 2
        : a.shape.top + a.shape.height / 2;
      const bCenter = horizontal
        ? b.shape.left + b.shape.width / 2
        : b.shape.top + b.shape.height / 2;
      return aCenter - bCenter || a.index - b.index;
    });
  const first = sorted[0].shape;
  const last = sorted[sorted.length - 1].shape;
  const start = horizontal ? first.left : first.top;
  const end = horizontal ? last.left + last.width : last.top + last.height;
  const totalSize = sorted.reduce(
    (sum, item) => sum + (horizontal ? item.shape.width : item.shape.height),
    0,
  );
  const gap = (end - start - totalSize) / (sorted.length - 1);
  const positions = new Map<string, ShapePosition>();
  let cursor = start;

  sorted.forEach(({ shape }) => {
    positions.set(shape.id, {
      id: shape.id,
      left: horizontal ? cursor : shape.left,
      top: horizontal ? shape.top : cursor,
    });
    cursor += (horizontal ? shape.width : shape.height) + gap;
  });

  return shapes.map((shape) => positions.get(shape.id) as ShapePosition);
}
```

- [x] **Step 4: Run tests and typecheck**

Run: `npm test`

Expected: 8 tests PASS.

Run: `npm run typecheck`

Expected: exit code 0.

- [x] **Step 5: Commit when Git metadata exists**

```bash
git add src/taskpane/alignment.ts tests/alignment.test.js
git commit -m "feat: add equal-gap distribution"
```

---

### Task 3: Matrix and Circle Layouts

**Files:**
- Modify: `src/taskpane/alignment.ts`
- Modify: `tests/alignment.test.js`

**Interfaces:**
- Consumes: `ShapeBounds`, `ShapePosition`, `requireShapes`, and `selectionBounds` from Task 1.
- Produces: `MatrixOptions`, `CircleOptions`, `arrangeMatrix(shapes, options)`, `arrangeCircle(shapes, options)`.

- [x] **Step 1: Add failing matrix tests**

Add `arrangeMatrix` to test imports and append:

```js
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
```

- [x] **Step 2: Add failing circle tests**

Add `arrangeCircle` to test imports and append:

```js
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
```

- [x] **Step 3: Run tests and confirm missing export failures**

Run: `npm test`

Expected: FAIL because matrix and circle functions are undefined.

- [x] **Step 4: Implement stable reading-order helper and matrix layout**

Append to `src/taskpane/alignment.ts`:

```ts
export interface MatrixOptions {
  columns: number;
  horizontalGap: number;
  verticalGap: number;
}

function readingOrder(shapes: ShapeBounds[]): ShapeBounds[] {
  return shapes
    .map((shape, index) => ({ shape, index }))
    .sort((a, b) => a.shape.top - b.shape.top || a.shape.left - b.shape.left || a.index - b.index)
    .map(({ shape }) => shape);
}

export function arrangeMatrix(shapes: ShapeBounds[], options: MatrixOptions): ShapePosition[] {
  requireShapes(shapes, 2);
  if (!Number.isInteger(options.columns) || options.columns < 1) {
    throw new Error("Columns must be a positive integer.");
  }
  if (options.columns > shapes.length) {
    throw new Error("Columns cannot exceed selected shape count.");
  }
  if (!Number.isFinite(options.horizontalGap) || !Number.isFinite(options.verticalGap)) {
    throw new Error("Matrix gaps must be valid numbers.");
  }

  const ordered = readingOrder(shapes);
  const rows = Math.ceil(ordered.length / options.columns);
  const columnWidths = Array.from({ length: options.columns }, () => 0);
  const rowHeights = Array.from({ length: rows }, () => 0);
  ordered.forEach((shape, index) => {
    const column = index % options.columns;
    const row = Math.floor(index / options.columns);
    columnWidths[column] = Math.max(columnWidths[column], shape.width);
    rowHeights[row] = Math.max(rowHeights[row], shape.height);
  });

  const bounds = selectionBounds(shapes);
  const columnLefts = columnWidths.map((_, column) =>
    bounds.left + columnWidths.slice(0, column).reduce((sum, width) => sum + width, 0) + column * options.horizontalGap,
  );
  const rowTops = rowHeights.map((_, row) =>
    bounds.top + rowHeights.slice(0, row).reduce((sum, height) => sum + height, 0) + row * options.verticalGap,
  );
  const positions = new Map<string, ShapePosition>();

  ordered.forEach((shape, index) => {
    const column = index % options.columns;
    const row = Math.floor(index / options.columns);
    positions.set(shape.id, {
      id: shape.id,
      left: columnLefts[column] + (columnWidths[column] - shape.width) / 2,
      top: rowTops[row] + (rowHeights[row] - shape.height) / 2,
    });
  });

  return shapes.map((shape) => positions.get(shape.id) as ShapePosition);
}
```

- [x] **Step 5: Implement circle layout**

Append to `src/taskpane/alignment.ts`:

```ts
export interface CircleOptions {
  radius: number;
  startAngle: number;
  clockwise: boolean;
}

function cleanCoordinate(value: number): number {
  return Math.abs(value) < 1e-10 ? 0 : Number(value.toFixed(10));
}

export function arrangeCircle(shapes: ShapeBounds[], options: CircleOptions): ShapePosition[] {
  requireShapes(shapes, 2);
  if (!Number.isFinite(options.radius) || options.radius <= 0) {
    throw new Error("Radius must be greater than 0.");
  }
  if (!Number.isFinite(options.startAngle)) {
    throw new Error("Start angle must be a valid number.");
  }

  const ordered = readingOrder(shapes);
  const bounds = selectionBounds(shapes);
  const centerX = (bounds.left + bounds.right) / 2;
  const centerY = (bounds.top + bounds.bottom) / 2;
  const direction = options.clockwise ? 1 : -1;
  const positions = new Map<string, ShapePosition>();

  ordered.forEach((shape, index) => {
    const degrees = options.startAngle + direction * (360 * index) / ordered.length;
    const radians = (degrees * Math.PI) / 180;
    positions.set(shape.id, {
      id: shape.id,
      left: cleanCoordinate(centerX + options.radius * Math.cos(radians) - shape.width / 2),
      top: cleanCoordinate(centerY + options.radius * Math.sin(radians) - shape.height / 2),
    });
  });

  return shapes.map((shape) => positions.get(shape.id) as ShapePosition);
}
```

- [x] **Step 6: Run all geometry tests and typecheck**

Run: `npm test`

Expected: 13 tests PASS.

Run: `npm run typecheck`

Expected: exit code 0.

- [x] **Step 7: Commit when Git metadata exists**

```bash
git add src/taskpane/alignment.ts tests/alignment.test.js
git commit -m "feat: add matrix and circle layouts"
```

---

### Task 4: PowerPoint Selection Adapter

**Files:**
- Modify: `src/taskpane/powerpoint.ts`

**Interfaces:**
- Consumes: `ShapeBounds`, `ShapePosition` from Task 1.
- Produces: `LayoutComputer`, `applyLayout(compute): Promise<number>`, `SmartAlignmentError`.

- [x] **Step 1: Replace starter action with typed Office adapter**

Replace `src/taskpane/powerpoint.ts` with:

```ts
import { ShapeBounds, ShapePosition } from "./alignment";

/* global Office, PowerPoint */

export type LayoutComputer = (shapes: ShapeBounds[]) => ShapePosition[];

export class SmartAlignmentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SmartAlignmentError";
  }
}

function assertSupported(): void {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    throw new SmartAlignmentError("Smart Alignment requires PowerPoint API 1.5.");
  }
}

export async function applyLayout(compute: LayoutComputer): Promise<number> {
  assertSupported();

  try {
    return await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedShapes();
      selected.load("items");
      await context.sync();

      selected.items.forEach((shape) => shape.load("id,left,top,width,height"));
      await context.sync();

      const bounds: ShapeBounds[] = selected.items.map((shape) => ({
        id: shape.id,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }));
      const positions = compute(bounds);
      const positionById = new Map(positions.map((position) => [position.id, position]));

      selected.items.forEach((shape) => {
        const position = positionById.get(shape.id);
        if (!position) {
          throw new SmartAlignmentError(`Layout returned no position for shape ${shape.id}.`);
        }
        shape.left = position.left;
        shape.top = position.top;
      });

      await context.sync();
      return selected.items.length;
    });
  } catch (error) {
    if (error instanceof SmartAlignmentError || error instanceof Error) {
      throw error;
    }
    throw new SmartAlignmentError("PowerPoint could not update selected shapes.");
  }
}
```

- [x] **Step 2: Typecheck Office adapter**

Run: `npm run typecheck`

Expected: exit code 0 with Office.js types resolving `getSelectedShapes`, shape coordinates, and `PowerPoint.run`.

- [x] **Step 3: Run production build**

Run: `npm run build`

Expected: Webpack exits 0 and writes task-pane assets under `dist/`.

- [x] **Step 4: Commit when Git metadata exists**

```bash
git add src/taskpane/powerpoint.ts
git commit -m "feat: connect layouts to PowerPoint selection"
```

---

### Task 5: Smart Alignment Task Pane

**Files:**
- Modify: `src/taskpane/taskpane.html`
- Modify: `src/taskpane/taskpane.css`
- Modify: `src/taskpane/taskpane.ts`
- Modify: `webpack.config.js`

**Interfaces:**
- Consumes: `alignShapes`, `distributeShapes`, `arrangeMatrix`, `arrangeCircle`, related option types, and `applyLayout`.
- Produces: user-facing controls for all ten operations and visible operation status.

- [x] **Step 1: Replace starter HTML with accessible controls**

Keep Office.js script and stylesheet links in `<head>`. Replace `<body>` with:

```html
<body class="ms-Fabric">
  <section id="sideload-msg" class="startup-message">
    Open Slideware inside PowerPoint to use Smart Alignment.
  </section>
  <main id="app-body" class="app" hidden>
    <header class="app-header">
      <p class="eyebrow">PRODUCTIVITY</p>
      <h1>Smart Alignment</h1>
      <p class="intro">Select shapes, choose an alignment or layout, then apply.</p>
    </header>

    <section class="panel" aria-labelledby="align-heading">
      <h2 id="align-heading">Align</h2>
      <div class="button-grid button-grid-3">
        <button type="button" data-align="left">Left</button>
        <button type="button" data-align="center">Center</button>
        <button type="button" data-align="right">Right</button>
        <button type="button" data-align="top">Top</button>
        <button type="button" data-align="middle">Middle</button>
        <button type="button" data-align="bottom">Bottom</button>
      </div>
    </section>

    <section class="panel" aria-labelledby="distribute-heading">
      <h2 id="distribute-heading">Distribute</h2>
      <div class="button-grid button-grid-2">
        <button type="button" data-distribute="horizontal">Horizontal</button>
        <button type="button" data-distribute="vertical">Vertical</button>
      </div>
    </section>

    <section class="panel" aria-labelledby="matrix-heading">
      <h2 id="matrix-heading">Matrix</h2>
      <div class="field-grid">
        <label>Columns <input id="matrix-columns" type="number" min="1" step="1" value="3"></label>
        <label>Column gap <input id="matrix-column-gap" type="number" step="1" value="16"></label>
        <label>Row gap <input id="matrix-row-gap" type="number" step="1" value="16"></label>
      </div>
      <button id="apply-matrix" class="primary" type="button">Arrange matrix</button>
    </section>

    <section class="panel" aria-labelledby="circle-heading">
      <h2 id="circle-heading">Circle</h2>
      <div class="field-grid">
        <label>Radius <input id="circle-radius" type="number" min="1" step="1" value="120"></label>
        <label>Start angle <input id="circle-angle" type="number" step="1" value="-90"></label>
        <label>Direction
          <select id="circle-direction">
            <option value="clockwise">Clockwise</option>
            <option value="counterclockwise">Counterclockwise</option>
          </select>
        </label>
      </div>
      <button id="apply-circle" class="primary" type="button">Arrange circle</button>
    </section>

    <p id="status" class="status" role="status" aria-live="polite">Ready.</p>
  </main>
</body>
```

- [x] **Step 2: Replace starter controller with operation dispatch**

Replace `src/taskpane/taskpane.ts` with:

```ts
import {
  AlignMode,
  DistributionAxis,
  alignShapes,
  arrangeCircle,
  arrangeMatrix,
  distributeShapes,
} from "./alignment";
import { applyLayout } from "./powerpoint";

/* global document, Office */

function requiredElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing task-pane element: ${id}.`);
  return element as T;
}

function numberValue(id: string): number {
  const value = Number(requiredElement<HTMLInputElement>(id).value);
  if (!Number.isFinite(value)) throw new Error(`${id.replace(/-/g, " ")} must be a valid number.`);
  return value;
}

function setBusy(busy: boolean): void {
  document.querySelectorAll<HTMLButtonElement>("button").forEach((button) => {
    button.disabled = busy;
  });
  document.body.setAttribute("aria-busy", String(busy));
}

function messageFor(error: unknown): string {
  return error instanceof Error ? error.message : "PowerPoint could not update selected shapes.";
}

async function execute(label: string, operation: Parameters<typeof applyLayout>[0]): Promise<void> {
  const status = requiredElement<HTMLElement>("status");
  setBusy(true);
  status.className = "status";
  status.textContent = `${label} in progress...`;
  try {
    const count = await applyLayout(operation);
    status.className = "status status-success";
    status.textContent = `${label} applied to ${count} shapes.`;
  } catch (error) {
    status.className = "status status-error";
    status.textContent = messageFor(error);
  } finally {
    setBusy(false);
  }
}

function bindControls(): void {
  document.querySelectorAll<HTMLButtonElement>("[data-align]").forEach((button) => {
    button.addEventListener("click", () => {
      const mode = button.dataset.align as AlignMode;
      void execute(`Align ${button.textContent?.toLowerCase()}`, (shapes) => alignShapes(shapes, mode));
    });
  });

  document.querySelectorAll<HTMLButtonElement>("[data-distribute]").forEach((button) => {
    button.addEventListener("click", () => {
      const axis = button.dataset.distribute as DistributionAxis;
      void execute(`Distribute ${axis}`, (shapes) => distributeShapes(shapes, axis));
    });
  });

  requiredElement<HTMLButtonElement>("apply-matrix").addEventListener("click", () => {
    void execute("Matrix", (shapes) =>
      arrangeMatrix(shapes, {
        columns: numberValue("matrix-columns"),
        horizontalGap: numberValue("matrix-column-gap"),
        verticalGap: numberValue("matrix-row-gap"),
      }),
    );
  });

  requiredElement<HTMLButtonElement>("apply-circle").addEventListener("click", () => {
    void execute("Circle", (shapes) =>
      arrangeCircle(shapes, {
        radius: numberValue("circle-radius"),
        startAngle: numberValue("circle-angle"),
        clockwise: requiredElement<HTMLSelectElement>("circle-direction").value === "clockwise",
      }),
    );
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  requiredElement<HTMLElement>("sideload-msg").hidden = true;
  requiredElement<HTMLElement>("app-body").hidden = false;
  bindControls();
});
```

- [x] **Step 3: Replace starter CSS with compact task-pane styling**

Replace `src/taskpane/taskpane.css` with:

```css
:root {
  color: #20262e;
  background: #f5f7f8;
  font-family: "Segoe UI", Arial, sans-serif;
}

* { box-sizing: border-box; }
html, body { min-width: 260px; min-height: 100%; margin: 0; }
body { background: #f5f7f8; }
button, input, select { font: inherit; }

.startup-message { padding: 24px; color: #4b5563; }
.app { padding: 18px; }
.app[hidden], .startup-message[hidden] { display: none; }
.app-header { margin-bottom: 16px; }
.eyebrow { margin: 0 0 4px; color: #66707a; font-size: 11px; font-weight: 700; letter-spacing: 0.12em; }
h1 { margin: 0; font-size: 24px; line-height: 1.2; }
.intro { margin: 8px 0 0; color: #58616b; line-height: 1.4; }

.panel { margin-top: 12px; padding: 14px; border: 1px solid #d9dee3; border-radius: 10px; background: #fff; }
.panel h2 { margin: 0 0 12px; font-size: 15px; }
.button-grid { display: grid; gap: 8px; }
.button-grid-3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
.button-grid-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }

button { min-height: 36px; padding: 7px 9px; border: 1px solid #c8ced4; border-radius: 7px; background: #fff; color: #20262e; cursor: pointer; }
button:hover { border-color: #0f6cbd; background: #f0f6fc; }
button:focus-visible, input:focus-visible, select:focus-visible { outline: 2px solid #0f6cbd; outline-offset: 2px; }
button:disabled { cursor: wait; opacity: 0.55; }
button.primary { width: 100%; margin-top: 12px; border-color: #0f6cbd; background: #0f6cbd; color: #fff; }
button.primary:hover { background: #115ea3; }

.field-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
label { display: grid; gap: 5px; color: #4b5563; font-size: 12px; }
input, select { width: 100%; min-height: 34px; padding: 6px 8px; border: 1px solid #c8ced4; border-radius: 6px; background: #fff; color: #20262e; }
.status { min-height: 20px; margin: 14px 2px 0; color: #58616b; font-size: 12px; }
.status-success { color: #0b6a3c; }
.status-error { color: #b42318; }

@media (max-width: 300px) {
  .button-grid-3, .field-grid { grid-template-columns: 1fr; }
}
```

- [x] **Step 4: Configure Webpack to emit task-pane CSS**

Add this rule after the HTML rule in `webpack.config.js`:

```js
{
  test: /\.css$/,
  type: "asset/resource",
  generator: {
    filename: "[name][ext]",
  },
},
```

`taskpane.html` keeps `<link href="taskpane.css" rel="stylesheet" type="text/css" />`; `html-loader` resolves that file and Webpack emits it at `dist/taskpane.css`.

- [x] **Step 5: Run automated checks**

Run: `npm test`

Expected: 13 tests PASS.

Run: `npm run typecheck`

Expected: exit code 0.

Run: `npm run lint`

Expected: exit code 0.

Run: `npm run build`

Expected: Webpack exits 0.

- [x] **Step 6: Commit when Git metadata exists**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.css src/taskpane/taskpane.ts webpack.config.js
git commit -m "feat: add Smart Alignment task pane"
```

---

### Task 6: Manifest and PowerPoint Acceptance Verification

**Files:**
- Modify: `manifest.json`
- Modify: `webpack.config.js`

**Interfaces:**
- Consumes: built task pane from Tasks 1 through 5.
- Produces: Slideware-branded ribbon entry and validated deployment manifest.

- [x] **Step 1: Replace remaining starter branding and remove unrelated command**

Change manifest values:

```json
"developer": {
  "name": "Slideware",
  "websiteUrl": "https://localhost:3000",
  "privacyUrl": "https://localhost:3000",
  "termsOfUseUrl": "https://localhost:3000"
},
"name": {
  "short": "Slideware",
  "full": "Slideware Smart Alignment"
},
"description": {
  "short": "Align and arrange selected PowerPoint shapes.",
  "full": "Align, distribute, and arrange selected PowerPoint shapes in matrix or circle layouts."
}
```

Rename ribbon group label to `Slideware`, pane button label to `Smart Alignment`, and matching supertip title to `Smart Alignment`. Set supertip description to `Align and arrange selected shapes.` Remove `ActionButton` from `controls`, remove its action from `CommandsRuntime`, then remove `CommandsRuntime` if no actions remain. Keep task-pane runtime and `Document.ReadWrite.User` permission.

- [x] **Step 2: Remove unused command entry only after manifest references are gone**

Delete `commands` from Webpack `entry`, remove commands `HtmlWebpackPlugin`, and leave source files untouched unless no other tooling references them. This removes unused build output without deleting user files.

- [x] **Step 3: Validate full package**

Run: `npm test`

Expected: 13 tests PASS.

Run: `npm run typecheck`

Expected: exit code 0.

Run: `npm run lint`

Expected: exit code 0.

Run: `npm run build`

Expected: Webpack exits 0.

Run: `npm run validate`

Expected: manifest validation exits 0.

- [ ] **Step 4: Run manual PowerPoint acceptance checks**

Run: `npm run start:desktop:powerpoint`

Verify in PowerPoint:

1. Select two mixed-size shapes. Run six alignment actions and confirm correct selection-edge or centerline placement.
2. Select three mixed-size shapes. Run horizontal and vertical distribution and confirm outer positions stay fixed while gaps match.
3. Select five shapes in mixed starting positions. Arrange as three-column matrix and confirm stable reading order, 16-point gaps, centered mixed-size cells, and partial final row.
4. Arrange same five shapes as circle with radius 120, angle -90, then both directions. Confirm equal center angles and unchanged dimensions and rotation.
5. Try each operation with insufficient selection and invalid numeric fields. Confirm visible error and no partial movement.
6. Trigger operations repeatedly. Confirm controls disable during each request and recover after success or failure.

- [ ] **Step 5: Stop debug session**

Run: `npm run stop`

Expected: PowerPoint add-in debug session closes cleanly.

- [x] **Step 6: Commit when Git metadata exists**

```bash
git add manifest.json webpack.config.js
git commit -m "chore: finalize Smart Alignment add-in"
```
