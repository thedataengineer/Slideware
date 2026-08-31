# Smart Alignment Design

## Decision

Build Smart Alignment as a task-pane feature with a pure TypeScript geometry core and a thin PowerPoint adapter. This keeps layout calculations testable without Office runtime access and limits Office API calls to selection reads and coordinate writes.

## Scope

Version 1 includes:

- Align left, horizontal center, right, top, vertical middle, and bottom.
- Distribute horizontally and vertically.
- Arrange selected shapes into a configurable matrix.
- Arrange selected shapes around a configurable circle.
- Validate selection counts and numeric inputs.
- Report success and failure in the task pane.

Version 1 does not resize, rotate, group, or restyle shapes. It does not add custom undo history, remote services, authentication, telemetry, or other Productivity features.

## Architecture

### Geometry core

`src/taskpane/alignment.ts` owns deterministic calculations. It receives plain shape bounds and returns new positions. It has no DOM or Office.js dependency.

Each shape uses this model:

```ts
interface ShapeBounds {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

interface ShapePosition {
  id: string;
  left: number;
  top: number;
}
```

### PowerPoint adapter

`src/taskpane/powerpoint.ts` loads selected shapes through `context.presentation.getSelectedShapes()`, reads `id`, `left`, `top`, `width`, and `height`, calls geometry functions, then writes `left` and `top` in one Office batch.

Selection access requires PowerPoint API 1.5. Shape position reads and writes require API 1.4. Runtime checks reject unsupported clients before any mutation.

### Task-pane controller

`src/taskpane/taskpane.ts` binds controls, parses settings, disables repeated submissions during Office calls, and renders concise status messages. UI state stays local to the task pane.

### Presentation

`src/taskpane/taskpane.html` contains four control groups: align, distribute, matrix, and circle. `src/taskpane/taskpane.css` provides compact keyboard-accessible controls sized for PowerPoint's narrow task pane.

## Layout Rules

### Ordering

Matrix and circle layouts sort selected shapes by `top`, then by `left`. Equal coordinates retain stable input order.

### Alignment

Alignment uses the bounding rectangle of the full selection:

- Left and top place each shape on the corresponding selection edge.
- Right and bottom account for each shape's width or height.
- Horizontal center and vertical middle align shape centers to selection centerlines.

Minimum selection: two shapes.

### Distribution

Horizontal distribution sorts by horizontal center. It preserves first and last shape positions, then creates equal edge-to-edge gaps between interior shapes. Vertical distribution follows the same rule using vertical centers and heights.

If total shape size exceeds available span, gaps become negative and shapes overlap evenly. This matches deterministic distribution behavior without silently resizing or expanding the selection.

Minimum selection: three shapes.

### Matrix

User inputs:

- Columns: positive integer, default `3`.
- Horizontal gap: number in points, default `16`.
- Vertical gap: number in points, default `16`.

Matrix anchors at selection's top-left corner. Each column width equals widest shape assigned to that column. Each row height equals tallest shape assigned to that row. Shapes center within their cells. Gaps separate cell boundaries. Extra shapes fill final partial row.

Minimum selection: two shapes. Column count cannot exceed selected shape count.

### Circle

User inputs:

- Radius: positive number in points, default `120`.
- Start angle: number in degrees, default `-90`, placing first shape at top.
- Direction: clockwise by default, switchable to counterclockwise.

Circle center equals center of current selection bounds. Radius measures from circle center to each shape center. Shapes receive equal angular spacing. Shape size and rotation remain unchanged.

Minimum selection: two shapes.

## Error Handling

Validation runs before coordinate mutation. Task pane reports:

- Unsupported PowerPoint API version.
- Insufficient selected shapes.
- Invalid columns, gaps, radius, or angle.
- Office runtime errors with user-safe messages.

Buttons remain disabled during an operation to prevent overlapping Office batches. Failed operations leave controls available for retry.

## Verification

Add unit tests for pure geometry functions covering mixed shape sizes, negative coordinates, overlapping distribution, incomplete matrix rows, circle direction, start angle, and invalid settings.

Run TypeScript compilation, lint, production build, and manifest validation. Manual PowerPoint verification covers selection reads, all ten actions, mixed-size shapes, repeated operations, and unsupported or insufficient selections.

## Acceptance Criteria

- Every alignment action changes selected shape coordinates according to selection bounds.
- Distribution preserves outer shapes and spaces interior shapes evenly.
- Matrix honors ordering, columns, gaps, mixed dimensions, and partial rows.
- Circle honors radius, start angle, direction, and selection-center anchoring.
- No operation changes shape width, height, rotation, styling, or grouping.
- Invalid input or selection produces visible feedback without partial coordinate writes.
- Production build and automated geometry tests pass.
