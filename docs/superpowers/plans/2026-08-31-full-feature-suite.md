# Full Feature Suite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship every menu feature: SmartBar, Shortcuts, Smart Selection, AI Editing, Checker, Agenda Maker, Templates, Branding Automation, Fonts & Colors, Custom Automations, GenAI Edit/Create/Translate, Super Search, Darwin, Claude MCP.

**Architecture:** Pure logic per feature in `src/taskpane/features/`, one Office.js adapter (`powerpoint.ts`) with capability guards, central op dispatcher shared by UI, shortcuts, automations, and the MCP bridge. Companion Node MCP server in `mcp-server/` bridges Claude to the pane over a localhost WebSocket.

**Tech Stack:** TypeScript 5.4, Office.js PowerPoint API 1.4–1.6 guarded, Anthropic Messages API (browser, BYO key), node:test, Webpack 5, Node `ws` (mcp-server only).

**Spec:** `docs/superpowers/specs/2026-08-31-full-feature-suite-design.md`

## Global Constraints

- No backend, no telemetry. AI key from user, localStorage only.
- Pure modules import nothing from Office/DOM; adapter owns all Office.js.
- Every pure module gets node:test coverage; `npm test`, `npm run typecheck`, `npm run lint`, `npm run build` stay green after every task.
- Commit after each task; plain messages, no AI attribution.
- Capability guards: selection read 1.5, shape creation/format writes 1.4, `setSelectedShapes` 1.6; unsupported hosts get named-version errors.

---

### Task 1: Generalized adapter writes + SmartBar geometry
- Create `src/taskpane/features/smartbar.ts`: `SizeMode = "width"|"height"|"both"`; `matchSizes(shapes, mode)` (reference = first shape; returns `ShapeUpdate[]` `{id, width?, height?}`); `swapPositions(shapes)` (exactly 2, else throw `Select exactly 2 shapes.`).
- Modify `src/taskpane/powerpoint.ts`: `ShapeUpdate = {id, left?, top?, width?, height?}`; `applyLayout` accepts computers returning `ShapeUpdate[]`, writes only present fields.
- Tests `tests/smartbar.test.js`: width/height/both mapping from mixed sizes; swap swaps left/top and keeps sizes; 1-shape and 3-shape swap throws.
- [x] Steps: failing tests → red → implement → green → typecheck → commit.

### Task 2: Tabbed UI shell + dispatcher + SmartBar UI
- Create `src/taskpane/dispatcher.ts`: `registerOp(id, {label, run, recordable?, params?})`, `dispatch(id, params?)`, `listOps()`; ops wrap existing execute/busy/status pattern.
- Modify `taskpane.html/css/ts`: SmartBar pinned row (6 align icons + W/H/S/Swap), tab strip Productivity | Branding | Gen AI with `[data-tab]` buttons + `[data-panel]` sections; existing alignment panels move into Productivity.
- Ops registered: `align.*`, `distribute.*`, `matrix`, `circle`, `size.width|height|both`, `swap`.
- Verify: build, lint, typecheck, tests. Commit.

### Task 3: Shortcuts
- Modify `taskpane.ts`: document keydown; ignore when target is input/select/textarea or busy; map L C R T M B → aligns, H V → distributes, W → size.width, E → size.height, S → swap.
- Cheat sheet panel in Productivity listing the keys.
- Commit.

### Task 4: Adapter deck/selection services
- Modify `powerpoint.ts`, produce:
  - `snapshotDeck(): Promise<DeckSnapshot>` (slides→shapes: id,name,type,left,top,width,height,text,fillColor,fontName,fontSize; title heuristic in pure helper),
  - `readSelection(): Promise<SnapshotShape[]>`,
  - `writeShapeFormats(updates: {id, fontName?, fontSize?, fontColor?, fillColor?}[], scope: "selection"|"deck")`,
  - `insertShapes(specs: InsertShapeSpec[])` (`textbox`/`rect`, API 1.4 guard),
  - `setSelection(shapeIds)` (1.6 guard), `gotoSlide(index)` (1.5), `replaceShapeText(id, text)`.
- Shared title heuristic `deriveTitle(shapes)` in new pure `src/taskpane/features/snapshot.ts` with tests (`tests/snapshot.test.js`): name containing "title" wins; else topmost text shape; else undefined.
- Commit.

### Task 5: Smart Selection
- Create `features/selection.ts`: `SelectionCriteria = {sameType?, sameFill?, sameSize?}`; `matchShapes(all, anchor, criteria)` → ids incl. anchor; size tolerance 1pt; no criteria → throw `Pick at least one criteria.`
- Tests: type match, fill match (case-insensitive hex), size tolerance boundary, combined criteria intersect, no-criteria throw.
- UI panel (Productivity): 3 checkboxes + Select button; unsupported 1.6 → status lists matched names instead.
- Commit.

### Task 6: Checker
- Create `features/checker.ts`: `audit(deck, slideSize={width:960,height:540})` → `Finding[]`; rules: `off-slide`, `tiny-font` (<12), `font-sprawl` (>3 families, one deck-level finding), `empty-text` (type textbox-ish + empty text), `overlong` (>300 chars).
- Tests: each rule positive+negative, sprawl counts distinct families, findings carry slideIndex/shapeId.
- UI panel: Run Checker → findings list (rule, slide n, shape name); item click → `gotoSlide`.
- Commit.

### Task 7: Agenda Maker
- Create `features/agenda.ts`: `buildAgenda(titles)` → `{text, lines}` numbered, skips empties; <1 title throws `No slide titles found.`
- Tests: numbering, empty-title skip, throw case.
- UI: Insert Agenda button → snapshot → titles (skip slide 1) → insert brand-styled textbox (left 60, top 120, width 600) on current slide.
- Commit.

### Task 8: Fonts & Colors (brand store)
- Create `features/branding.ts`: `Brand {headingFont, bodyFont, colors: string[6]}`, `defaultBrand`, `parseBrand(json)` (invalid → default), `serializeBrand`, `normalizeHex` (#RGB→#RRGGBB, invalid throws).
- Tests: parse round-trip, corrupt JSON → default, hex normalize + invalid throw.
- UI (Branding tab): font inputs, 6 color inputs, Save (localStorage `slideware.brand`), swatch row with Fill/Text mode toggle → `writeShapeFormats` on selection.
- Commit.

### Task 9: Branding Automation
- `features/branding.ts` add: `brandSelectionFormats(shapes, brand)` (fontName=bodyFont, fontColor=colors[0]) and `brandDeckFormats(deck, brand)` (all text shapes → fontName).
- Tests: format arrays built only for text-bearing shapes; colors mapped.
- UI: Apply Brand to Selection / Fix Fonts Across Deck buttons; ops `brand.applySelection`, `brand.applyDeck`.
- Commit.

### Task 10: Templates
- Create `features/templates.ts`: `TemplateName = "title-block"|"kpi-row"|"quote-card"|"section-divider"`; `templateShapes(name, brand, slideSize)` → `InsertShapeSpec[]` (positions/sizes derived from slideSize, fonts/colors from brand); unknown name throws.
- Tests: each template shape counts + within-slide bounds + brand fonts applied; unknown name throws.
- UI: 4-button grid → `insertShapes`.
- Commit.

### Task 11: Custom Automations
- Create `features/automations.ts`: `Recorder` (start/stop/recordStep, ignores non-recordable), `Automation {name, steps}`, `parseAutomations`/`serializeAutomations`, `validateName` (non-empty, unique).
- Tests: record sequence, non-recordable skipped, serialize round-trip, duplicate name throws.
- Dispatcher: `recordable` ops feed active recorder; UI (Branding tab): Record/Stop, name+Save, saved list with Run (sequential dispatch) / Delete. Storage `slideware.automations`.
- Commit.

### Task 12: AI client + prompts (read claude-api skill first)
- Create `src/taskpane/ai.ts`: `callClaude({apiKey, system, messages, maxTokens})` → text; fetch `https://api.anthropic.com/v1/messages`, headers incl. `anthropic-dangerous-direct-browser-access: true`; missing key/HTTP errors → typed messages.
- Create `features/prompts.ts` (pure): `editPrompt(text, instruction)`, `presetPrompt(text, preset)` (proofread/shorten/expand/clarify), `createPrompt(topic)` (strict JSON contract `{title,bullets}` + `parseCreateResponse`), `translatePrompt(text, language)`, `darwinSystem(outline)` (truncate 8000 chars).
- Tests (prompts only): preset wording per mode, create parser accepts fenced/bare JSON + throws on garbage, outline truncation.
- Gen AI tab: API key field (password, localStorage `slideware.apiKey`), model constant.
- Commit.

### Task 13: AI Editing + GenAI Edit / Create / Translate UIs
- Productivity: 4 preset buttons → anchor selected shape text → callClaude → `replaceShapeText`.
- Gen AI: Edit textarea+button (freeform); Create textarea+button → parse JSON → insert title+body textboxes (brand fonts); Translate language select+button → per selected text shape replace.
- Commit.

### Task 14: Super Search
- Create `features/search.ts`: `searchDeck(deck, query)` → `Hit {slideIndex, shapeId, shapeName, snippet}` (±30 chars context, case-insensitive; blank query throws).
- Tests: match, snippet windowing, multi-hit ordering by slide, blank throw.
- Gen AI tab: search box + results; click → gotoSlide + guarded setSelection.
- Commit.

### Task 15: Darwin chat
- Gen AI tab: chat log + input; system = `darwinSystem(outline from snapshotDeck)`; history in-memory; renders user/assistant bubbles; errors to status.
- Commit.

### Task 16: Claude MCP companion
- Create `mcp-server/protocol.js` (pure): `handleMessage(state, msg)` → `{responses[], effects[]}` for initialize/tools\/list/tools\/call/ping; unknown method → JSON-RPC -32601; tool catalog with input schemas.
- Create `mcp-server/server.js`: stdio newline-JSON loop + `ws` server 127.0.0.1:3711; tools/call → pending map `{callId}` → WS → response or 30s timeout / no-pane error.
- Create `mcp-server/package.json` (private, dep ws), `mcp-server/README.md` (Claude Desktop + `claude mcp add` snippets).
- Create `src/taskpane/bridge.ts`: connect/disconnect WS client, on message dispatch op via dispatcher + reply; Gen AI tab Connect toggle + status dot.
- Tests `tests/mcp-protocol.test.js`: initialize handshake shape, tools/list catalog, tools/call routes effect, unknown method error.
- Commit.

### Task 17: Final verification + docs
- Run: npm test, typecheck, lint, build, validate — all green.
- Update `docs/superpowers/specs` if drifted; check plan boxes; commit.
- Manual PowerPoint acceptance checklist (user): every panel operates on a real deck; MCP: `node mcp-server/server.js` + Claude Desktop config + Connect toggle.

---

## Status (2026-08-31)

Tasks 1-16 complete and committed; every check green (71 tests, typecheck, lint, webpack build, manifest validate). MCP stdio handshake and WS bridge round-trip verified with a simulated pane. Remaining: manual PowerPoint acceptance (Task 17 checklist) — run `npm run start:desktop:powerpoint`, exercise each panel on a real deck, and for MCP: `cd mcp-server && npm install && node server.js`, register with Claude per mcp-server/README.md, press Connect in the Gen AI tab.
