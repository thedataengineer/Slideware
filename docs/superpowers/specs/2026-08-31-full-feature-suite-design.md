# Slideware Full Feature Suite Design

Date: 2026-08-31
Status: Approved for implementation (autonomous session; user mandate: "Finish all the features")

## Goal

Extend the Slideware PowerPoint add-in from Smart Alignment alone to the full product menu:

- PRODUCTIVITY: SmartBar, Shortcuts, Smart Alignment (done), Smart Selection, AI Editing, Checker, Agenda Maker
- BRANDING: Templates, Branding Automation, Fonts & Colors, Custom Automations
- GEN AI: Edit, Create, Translate, Super Search, Darwin, Claude MCP

## Constraints

- Office.js PowerPoint add-in, task pane only. No backend service.
- [ASSUMPTION] Gen AI features call the Anthropic Messages API directly from the task pane with a user-supplied API key (browser direct access header). Key stored in localStorage, never bundled.
- Capability guards per Office API requirement set: base 1.5 (selection read), shape creation and text/format writes 1.4+, programmatic shape selection 1.6, guarded with `isSetSupported` and clear errors.
- Existing shape dimensions, rotation, grouping stay untouched unless the operation's purpose is formatting.
- Pure logic stays in testable modules; Office.js touches live only in the adapter.
- No AI attribution in commits.

## Architecture

```
src/taskpane/
  alignment.ts        existing pure geometry (unchanged)
  features/
    smartbar.ts       pure: sameSize / swapPositions position+size computations
    selection.ts      pure: matchShapes(snapshotShapes, criteria) -> ids
    checker.ts        pure: audit(DeckSnapshot) -> Finding[]
    agenda.ts         pure: buildAgenda(slideTitles) -> text block spec
    templates.ts      pure: templateShapes(name, brand, slideSize) -> InsertShapeSpec[]
    branding.ts       pure: Brand type, defaults, load/save codec, format ops builder
    automations.ts    pure: step recorder/registry, (de)serialization, validation
    search.ts         pure: searchDeck(DeckSnapshot, query) -> Hit[]
    prompts.ts        pure: prompt builders for edit/create/translate/darwin
  ai.ts               Anthropic Messages API client (fetch)
  bridge.ts           WebSocket client for MCP companion, command router
  dispatcher.ts       named op registry: UI, shortcuts, automations, MCP all dispatch here
  powerpoint.ts       adapter: applyLayout (existing) + snapshotDeck, readSelection,
                      writeShapeFormats, insertShapes, setSelection, gotoSlide, replaceShapeText
  taskpane.ts         UI wiring: tabs, controls, busy/status, shortcut keydown
  taskpane.html/css   tabbed layout: SmartBar pinned, tabs Productivity | Branding | Gen AI
mcp-server/
  package.json        own package, dep: ws
  server.js           stdio MCP server (hand-rolled JSON-RPC 2.0, newline-delimited)
                      + WS bridge on 127.0.0.1:3711; forwards tools/call to connected pane
  protocol.js         pure framing/dispatch (unit tested with node:test)
  README.md           Claude Desktop / Claude Code config snippet
tests/                node:test suites per pure module
```

### Data contracts

- `SnapshotShape`: `{ id, name, type, geometricType?, left, top, width, height, text?, fillColor?, fontName?, fontSize?, fontColor? }`
- `DeckSnapshot`: `{ slideCount, slides: [{ id, index, title?, shapes: SnapshotShape[] }] }`. Title heuristic: first shape whose name contains "Title", else topmost text shape.
- `Brand`: `{ headingFont, bodyFont, colors: string[6] }` (hex). Default: Segoe UI / Segoe UI, neutral palette.
- `AutomationStep`: `{ op: string, params?: Record<string, unknown> }`. Automation: `{ name, steps[] }`, stored as JSON in localStorage.
- `Finding` (checker): `{ rule, slideIndex, shapeId?, message }`.
- Dispatcher op ids: `align.left|center|right|top|middle|bottom`, `distribute.horizontal|vertical`, `matrix`, `circle`, `size.width|height|both`, `swap`, `select.smart`, `brand.applySelection`, `brand.applyDeck`, `template.insert`, `agenda.insert`, `checker.run`, `search`, `ai.edit|create|translate|preset`, plus MCP-only reads.

## Feature designs

### SmartBar (Productivity)
Pinned button row above tabs. Ops: 6 aligns (reuse alignShapes), Same Width / Same Height / Same Size (reference = first shape in selection order; positions unchanged, plan writes width/height), Swap (exactly 2 shapes, swap left/top). Pure math in `smartbar.ts`.

### Shortcuts
`keydown` on document, active when focus is not in an input/select/textarea. Map: L/C/R/T/M/B -> aligns, H/V -> distribute, W -> same width, E -> same height, S -> swap. Cheat sheet rendered in Productivity tab. No manifest-level shortcuts (extendedOverrides not broadly available for PowerPoint). Single-letter keys acceptable because pane focus is deliberate.

### Smart Selection
Reads current slide snapshot + one anchor shape (current selection's first shape). Criteria checkboxes: same shape type, same fill color, same size (±1pt tolerance). `matchShapes` returns matching ids on that slide; adapter calls `setSelectedShapes` (API 1.6, guarded; on unsupported hosts shows the matched count + names instead).

### AI Editing (Productivity) and GenAI Edit
Same engine. AI Editing = preset buttons (Proofread, Shorten, Expand, Clarify); GenAI Edit = freeform instruction textarea. Flow: read selected shapes' text (first shape with text), build prompt (`prompts.ts`), call `ai.ts`, write result back via `replaceShapeText`. Multi-shape: apply per shape independently for Translate; single anchor shape for Edit presets/freeform.

### Checker
`snapshotDeck` then pure rules: (1) shape outside slide bounds (slide size read from presentation; fallback 960x540pt), (2) font size < 12pt, (3) > 3 distinct font families deck-wide, (4) empty text boxes, (5) text blocks > 300 chars (overlong bullet heuristic). Results list: rule label, slide number, shape name. Click focuses slide via `gotoSlide` when supported.

### Agenda Maker
`snapshotDeck` -> titles list (skip slide 1). `buildAgenda` returns numbered lines + layout spec. Insert as text box on the CURRENT slide at brand-styled position (left 60, top 120, width 600). [ASSUMPTION] Creating a new slide programmatically needs API 1.8; v1 inserts on current slide, doc'd in UI copy.

### Templates (Branding)
Named templates: Title Block, KPI Row (3 tiles), Quote Card, Section Divider. `templateShapes(name, brand, slideSize)` emits InsertShapeSpec[]: `{ kind: "textbox"|"rect", left, top, width, height, text?, fillColor?, fontName?, fontSize?, fontColor? }`. Adapter inserts via `shapes.addTextBox` / `addGeometricShape` (API 1.4 guarded), applies fill/font.

### Branding Automation
Apply Brand to Selection: set fontName (body font), font color (palette[0] text color), fill for rect-like shapes (palette accents untouched — only font normalization plus optional fill toggle). Apply Fonts Deck-wide: iterate all slides' shapes with text, set fontName to body font. Explicit per-op formats built in `branding.ts` so tests cover mapping.

### Fonts & Colors
Brand editor form: heading font, body font, 6 color inputs. Save -> localStorage (`slideware.brand`). Swatch click -> apply as fill (Fill mode) or font color (Text mode) to selected shapes.

### Custom Automations
Dispatcher instrumented: when "recording", each parameterizable op appends an AutomationStep. Controls: Record/Stop, name field, Save; saved list with Run/Delete. Replay = sequential dispatch with current selection. Steps requiring prompts (AI ops) excluded from recording (v1).

### Super Search
Search box -> `snapshotDeck` -> `searchDeck` (case-insensitive substring, returns slide index, shape id/name, snippet with match context). Result click -> `gotoSlide` + guarded shape select.

### GenAI Create
Prompt textarea -> Claude returns strict JSON `{ title, bullets[] }` (prompt enforces) -> insert title text box + body text box on current slide, brand-styled.

### Translate
Language dropdown (10 common languages + free text) -> per selected shape with text, translate text, write back preserving nothing but text (formatting retained since only text replaced).

### Darwin
Chat panel in Gen AI tab. System prompt: presentation coach persona + deck outline (titles + per-slide text truncated to keep payload sane, cap ~8k chars). History kept in memory (session), rendered as chat bubbles. Read-only advisor (no tool-use writes in v1).

### Claude MCP
Companion process, not part of the web bundle:
- `mcp-server/server.js`: MCP over stdio (JSON-RPC 2.0, newline-delimited): `initialize`, `tools/list`, `tools/call`, `ping`. Hand-rolled protocol in `protocol.js` (pure, tested).
- WS server 127.0.0.1:3711 (dep: `ws`). Pane connects via `bridge.ts` when user toggles Connect in Gen AI tab. `tools/call` -> `{ callId, op, params }` over WS -> pane dispatches -> `{ callId, ok, result|error }` back. 30s timeout, "no pane connected" error when socket absent.
- Tools exposed: `get_deck_outline`, `get_selected_shapes`, `align_shapes`, `distribute_shapes`, `arrange_matrix`, `arrange_circle`, `same_size`, `swap_shapes`, `smart_select`, `apply_branding`, `insert_template`, `insert_agenda`, `run_checker`, `search_deck`, `goto_slide`, `set_shape_text`.
- README: config snippet for Claude Desktop (`mcpServers`) and Claude Code (`claude mcp add`).
- [ASSUMPTION] ws dependency lives only in `mcp-server/package.json`; root package untouched.

## UI layout

- Header (existing) + SmartBar (pinned) + tab strip: Productivity | Branding | Gen AI.
- Productivity: Align/Distribute/Matrix/Circle (existing panels), Smart Selection, AI Editing presets, Checker, Agenda, Shortcuts cheat sheet.
- Branding: Fonts & Colors editor, swatches, Branding Automation buttons, Templates grid, Custom Automations.
- Gen AI: API key settings (password input, saved locally note), Edit, Create, Translate, Darwin chat, Claude MCP connect.
- Single `#status` live region retained. Busy state disables buttons.

## Error handling

- Every op through `execute()`-style wrapper: busy, status success/error, no partial writes where computable upfront (validate before write).
- AI errors: missing key -> "Add your Anthropic API key in Gen AI settings."; HTTP errors surfaced with status code; JSON-parse failures of Create -> retry-free explicit error.
- Capability errors name the required PowerPoint API version.

## Testing

- node:test per pure module: smartbar, selection, checker, agenda, templates, branding, automations, search, prompts, mcp protocol.
- tsconfig.test.json includes all pure feature modules (no Office/DOM imports there).
- Existing 18 alignment tests keep passing. Typecheck, ESLint, webpack build, manifest validate all green.
- Manual PowerPoint acceptance deferred to user (documented checklist in plan).

## Out of scope (v1)

- Manifest keyboard shortcuts (extendedOverrides), new-slide creation (API 1.8 hosts), AI tool-use writes from Darwin, template thumbnails, image/media templates, telemetry of any kind.
