import {
  AlignMode,
  DistributionAxis,
  alignShapes,
  arrangeCircle,
  arrangeMatrix,
  distributeShapes,
} from "./alignment";
import { SizeMode, matchSizes, swapPositions } from "./features/smartbar";
import { matchShapes } from "./features/selection";
import { audit } from "./features/checker";
import { buildAgenda } from "./features/agenda";
import {
  Brand,
  brandDeckFormats,
  brandSelectionFormats,
  defaultBrand,
  normalizeHex,
  parseBrand,
  serializeBrand,
} from "./features/branding";
import { templateShapes } from "./features/templates";
import {
  Automation,
  Recorder,
  parseAutomations,
  serializeAutomations,
  validateName,
} from "./features/automations";
import {
  createPrompt,
  darwinSystem,
  editPrompt,
  parseCreateResponse,
  presetPrompt,
  translatePrompt,
} from "./features/prompts";
import { searchDeck } from "./features/search";
import { callClaude } from "./ai";
import { connectBridge, disconnectBridge, isBridgeConnected, onBridgeStatus } from "./bridge";
import { dispatch, registerOp, setRecordListener } from "./dispatcher";
import {
  applyLayout,
  canSelectShapes,
  gotoSlide,
  insertShapes,
  readSelection,
  replaceShapeText,
  setSelection,
  snapshotDeck,
  writeShapeFormats,
} from "./powerpoint";

/* global localStorage */

/* global document, Office, HTMLElement, HTMLInputElement, HTMLButtonElement, HTMLSelectElement */

export function requiredElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing task-pane element: ${id}.`);
  return element as T;
}

function numberValue(id: string): number {
  const value = Number(requiredElement<HTMLInputElement>(id).value);
  if (!Number.isFinite(value)) throw new Error(`${id.replace(/-/g, " ")} must be a valid number.`);
  return value;
}

let busy = false;

function setBusy(nextBusy: boolean): void {
  busy = nextBusy;
  document.querySelectorAll<HTMLButtonElement>("button").forEach((button) => {
    button.disabled = nextBusy;
  });
  document.body.setAttribute("aria-busy", String(nextBusy));
}

export function isBusy(): boolean {
  return busy;
}

function messageFor(error: unknown): string {
  return error instanceof Error ? error.message : "PowerPoint could not complete the operation.";
}

export function showStatus(kind: "info" | "success" | "error", text: string): void {
  const status = requiredElement<HTMLElement>("status");
  const classes = {
    info: "status",
    success: "status status-success",
    error: "status status-error",
  };
  status.className = classes[kind];
  status.textContent = text;
}

export async function runOp(id: string, params?: Record<string, unknown>): Promise<void> {
  if (busy) return;
  setBusy(true);
  showStatus("info", "Working...");
  try {
    const message = await dispatch(id, params);
    showStatus("success", message);
  } catch (error) {
    showStatus("error", messageFor(error));
  } finally {
    setBusy(false);
  }
}

function registerLayoutOps(): void {
  const alignModes: AlignMode[] = ["left", "center", "right", "top", "middle", "bottom"];
  alignModes.forEach((mode) => {
    registerOp(`align.${mode}`, {
      label: `Align ${mode}`,
      recordable: true,
      run: async () => {
        const count = await applyLayout((shapes) => alignShapes(shapes, mode));
        return `Align ${mode} applied to ${count} shapes.`;
      },
    });
  });

  const axes: DistributionAxis[] = ["horizontal", "vertical"];
  axes.forEach((axis) => {
    registerOp(`distribute.${axis}`, {
      label: `Distribute ${axis}`,
      recordable: true,
      run: async () => {
        const count = await applyLayout((shapes) => distributeShapes(shapes, axis));
        return `Distribute ${axis} applied to ${count} shapes.`;
      },
    });
  });

  registerOp("matrix", {
    label: "Arrange matrix",
    recordable: true,
    run: async (params) => {
      const options = {
        columns: Number(params?.columns),
        horizontalGap: Number(params?.horizontalGap),
        verticalGap: Number(params?.verticalGap),
      };
      const count = await applyLayout((shapes) => arrangeMatrix(shapes, options));
      return `Matrix applied to ${count} shapes.`;
    },
  });

  registerOp("circle", {
    label: "Arrange circle",
    recordable: true,
    run: async (params) => {
      const options = {
        radius: Number(params?.radius),
        startAngle: Number(params?.startAngle),
        clockwise: params?.clockwise !== false,
      };
      const count = await applyLayout((shapes) => arrangeCircle(shapes, options));
      return `Circle applied to ${count} shapes.`;
    },
  });

  const sizeModes: SizeMode[] = ["width", "height", "both"];
  sizeModes.forEach((mode) => {
    registerOp(`size.${mode}`, {
      label: `Same ${mode === "both" ? "size" : mode}`,
      recordable: true,
      run: async () => {
        const count = await applyLayout((shapes) => matchSizes(shapes, mode));
        return `Matched ${mode === "both" ? "size" : mode} on ${count} shapes.`;
      },
    });
  });

  registerOp("swap", {
    label: "Swap positions",
    recordable: true,
    run: async () => {
      await applyLayout((shapes) => swapPositions(shapes));
      return "Swapped shape positions.";
    },
  });
}

function checkboxValue(id: string): boolean {
  return requiredElement<HTMLInputElement>(id).checked;
}

function registerSelectionOp(): void {
  registerOp("select.smart", {
    label: "Smart Selection",
    run: async (params) => {
      const criteria = {
        sameType: params?.sameType === true,
        sameFill: params?.sameFill === true,
        sameSize: params?.sameSize === true,
      };
      const selected = await readSelection();
      if (selected.length === 0) throw new Error("Select an anchor shape first.");
      const anchor = selected[0];

      const deck = await snapshotDeck();
      const slide = deck.slides.find((candidate) =>
        candidate.shapes.some((shape) => shape.id === anchor.id)
      );
      if (!slide) throw new Error("Could not locate the anchor shape's slide.");

      const ids = matchShapes(slide.shapes, anchor, criteria);
      if (canSelectShapes()) {
        await setSelection(ids);
        return `Selected ${ids.length} matching shapes.`;
      }
      const names = slide.shapes
        .filter((shape) => ids.includes(shape.id))
        .map((shape) => shape.name)
        .join(", ");
      return `Matched ${ids.length} shapes (selection needs API 1.6): ${names}`;
    },
  });
}

function bindSelectionControls(): void {
  requiredElement<HTMLButtonElement>("smart-select").addEventListener("click", () => {
    void runOp("select.smart", {
      sameType: checkboxValue("sel-type"),
      sameFill: checkboxValue("sel-fill"),
      sameSize: checkboxValue("sel-size"),
    });
  });
}

async function runTask(work: () => Promise<string>): Promise<void> {
  if (busy) return;
  setBusy(true);
  showStatus("info", "Working...");
  try {
    showStatus("success", await work());
  } catch (error) {
    showStatus("error", messageFor(error));
  } finally {
    setBusy(false);
  }
}

function clearList(id: string): HTMLElement {
  const list = requiredElement<HTMLElement>(id);
  list.textContent = "";
  return list;
}

function addListItem(list: HTMLElement, text: string, meta?: string, onClick?: () => void): void {
  const item = document.createElement("li");
  item.textContent = text;
  if (meta) {
    const metaLine = document.createElement("div");
    metaLine.className = "meta";
    metaLine.textContent = meta;
    item.appendChild(metaLine);
  }
  if (onClick) {
    item.className = "clickable";
    item.addEventListener("click", onClick);
  }
  list.appendChild(item);
}

function bindCheckerControls(): void {
  requiredElement<HTMLButtonElement>("run-checker").addEventListener("click", () => {
    void runTask(async () => {
      const deck = await snapshotDeck();
      const findings = audit(deck);
      const list = clearList("checker-results");
      findings.forEach((finding) => {
        const meta = finding.slideIndex
          ? `Slide ${finding.slideIndex} · ${finding.rule}`
          : finding.rule;
        const jump = finding.slideIndex
          ? () =>
              void runTask(async () => {
                await gotoSlide(finding.slideIndex as number);
                return `Moved to slide ${finding.slideIndex}.`;
              })
          : undefined;
        addListItem(list, finding.message, meta, jump);
      });
      return findings.length === 0
        ? "Checker found no issues."
        : `Checker found ${findings.length} issues.`;
    });
  });
}

const BRAND_STORAGE_KEY = "slideware.brand";

function loadBrand(): Brand {
  try {
    return parseBrand(localStorage.getItem(BRAND_STORAGE_KEY));
  } catch {
    return defaultBrand();
  }
}

function saveBrand(brand: Brand): void {
  try {
    localStorage.setItem(BRAND_STORAGE_KEY, serializeBrand(brand));
  } catch {
    // Storage unavailable; brand stays session-only.
  }
}

function brandFromForm(): Brand {
  const colors = [0, 1, 2, 3, 4, 5].map((index) =>
    normalizeHex(requiredElement<HTMLInputElement>(`brand-color-${index}`).value)
  );
  return {
    headingFont: requiredElement<HTMLInputElement>("brand-heading-font").value.trim() || "Segoe UI",
    bodyFont: requiredElement<HTMLInputElement>("brand-body-font").value.trim() || "Segoe UI",
    colors,
  };
}

function renderBrandForm(brand: Brand): void {
  requiredElement<HTMLInputElement>("brand-heading-font").value = brand.headingFont;
  requiredElement<HTMLInputElement>("brand-body-font").value = brand.bodyFont;
  brand.colors.forEach((color, index) => {
    requiredElement<HTMLInputElement>(`brand-color-${index}`).value = color;
  });
  renderSwatches(brand);
}

function renderSwatches(brand: Brand): void {
  const container = requiredElement<HTMLElement>("swatches");
  container.textContent = "";
  brand.colors.forEach((color) => {
    const swatch = document.createElement("button");
    swatch.type = "button";
    swatch.style.background = color;
    swatch.title = `Apply ${color}`;
    swatch.setAttribute("aria-label", `Apply ${color}`);
    swatch.addEventListener("click", () => {
      const mode = requiredElement<HTMLSelectElement>("swatch-mode").value;
      void runTask(async () => {
        const selected = await readSelection();
        if (selected.length === 0) throw new Error("Select at least 1 shape.");
        const formats = selected.map((shape) =>
          mode === "fill" ? { id: shape.id, fillColor: color } : { id: shape.id, fontColor: color }
        );
        const applied = await writeShapeFormats(formats);
        return `Applied ${color} as ${mode === "fill" ? "fill" : "text color"} to ${applied} shapes.`;
      });
    });
    container.appendChild(swatch);
  });
}

function registerBrandOps(): void {
  registerOp("brand.applySelection", {
    label: "Apply brand to selection",
    recordable: true,
    run: async () => {
      const selected = await readSelection();
      if (selected.length === 0) throw new Error("Select at least 1 shape.");
      const formats = brandSelectionFormats(selected, loadBrand());
      if (formats.length === 0) throw new Error("The selection has no text to brand.");
      const applied = await writeShapeFormats(formats);
      return `Branded ${applied} shapes.`;
    },
  });

  registerOp("brand.applyDeck", {
    label: "Fix fonts across deck",
    recordable: true,
    run: async () => {
      const deck = await snapshotDeck();
      const formats = brandDeckFormats(deck, loadBrand());
      if (formats.length === 0) throw new Error("The deck has no text shapes.");
      const applied = await writeShapeFormats(formats);
      return `Set the brand font on ${applied} shapes.`;
    },
  });

  registerOp("agenda.insert", {
    label: "Insert agenda",
    recordable: true,
    run: async () => {
      const deck = await snapshotDeck();
      const titles = deck.slides.slice(1).map((slide) => slide.title);
      const agenda = buildAgenda(titles);
      const brand = loadBrand();
      await insertShapes([
        {
          kind: "textbox",
          left: 60,
          top: 60,
          width: 600,
          height: 40,
          text: "Agenda",
          fontName: brand.headingFont,
          fontSize: 28,
          fontColor: brand.colors[0],
          bold: true,
        },
        {
          kind: "textbox",
          left: 60,
          top: 120,
          width: 600,
          height: 40 + agenda.lines.length * 24,
          text: agenda.text,
          fontName: brand.bodyFont,
          fontSize: 18,
          fontColor: brand.colors[0],
        },
      ]);
      return `Inserted an agenda with ${agenda.lines.length} items.`;
    },
  });
}

function bindBrandControls(): void {
  renderBrandForm(loadBrand());

  requiredElement<HTMLButtonElement>("save-brand").addEventListener("click", () => {
    void runTask(async () => {
      const brand = brandFromForm();
      saveBrand(brand);
      renderSwatches(brand);
      return "Brand saved.";
    });
  });

  requiredElement<HTMLButtonElement>("apply-brand-selection").addEventListener(
    "click",
    () => void runOp("brand.applySelection")
  );
  requiredElement<HTMLButtonElement>("apply-brand-deck").addEventListener(
    "click",
    () => void runOp("brand.applyDeck")
  );
  requiredElement<HTMLButtonElement>("insert-agenda").addEventListener(
    "click",
    () => void runOp("agenda.insert")
  );
}

const SLIDE_SIZE = { width: 960, height: 540 };
const AUTOMATIONS_STORAGE_KEY = "slideware.automations";
const recorder = new Recorder();
let pendingSteps: ReturnType<Recorder["stop"]> | null = null;

function registerTemplateOp(): void {
  registerOp("template.insert", {
    label: "Insert template",
    recordable: true,
    run: async (params) => {
      const name = String(params?.name ?? "");
      const specs = templateShapes(name, loadBrand(), SLIDE_SIZE);
      await insertShapes(specs);
      return `Inserted the ${name.replace(/-/g, " ")} template.`;
    },
  });
}

function bindTemplateControls(): void {
  document.querySelectorAll<HTMLButtonElement>("[data-template]").forEach((button) => {
    button.addEventListener(
      "click",
      () => void runOp("template.insert", { name: button.dataset.template })
    );
  });
}

function loadAutomations(): Automation[] {
  try {
    return parseAutomations(localStorage.getItem(AUTOMATIONS_STORAGE_KEY));
  } catch {
    return [];
  }
}

function saveAutomations(automations: Automation[]): void {
  try {
    localStorage.setItem(AUTOMATIONS_STORAGE_KEY, serializeAutomations(automations));
  } catch {
    // Storage unavailable; automations stay session-only.
  }
}

function renderAutomations(): void {
  const list = clearList("automation-list");
  loadAutomations().forEach((automation) => {
    const item = document.createElement("li");
    const title = document.createElement("div");
    title.textContent = automation.name;
    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = `${automation.steps.length} steps`;
    const actions = document.createElement("div");
    actions.className = "row-actions";

    const runButton = document.createElement("button");
    runButton.type = "button";
    runButton.textContent = "Run";
    runButton.addEventListener("click", () => {
      void runTask(async () => {
        for (const step of automation.steps) {
          await dispatch(step.op, step.params);
        }
        return `Ran "${automation.name}" (${automation.steps.length} steps).`;
      });
    });

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.textContent = "Delete";
    deleteButton.addEventListener("click", () => {
      saveAutomations(loadAutomations().filter((entry) => entry.name !== automation.name));
      renderAutomations();
      showStatus("success", `Deleted "${automation.name}".`);
    });

    actions.append(runButton, deleteButton);
    item.append(title, meta, actions);
    list.appendChild(item);
  });
}

function bindAutomationControls(): void {
  const recordButton = requiredElement<HTMLButtonElement>("record-automation");
  const saveButton = requiredElement<HTMLButtonElement>("save-automation");

  setRecordListener((id, params) => recorder.recordStep(id, params));

  recordButton.addEventListener("click", () => {
    if (recorder.isRecording()) {
      pendingSteps = recorder.stop();
      recordButton.textContent = "Record";
      recordButton.classList.remove("recording");
      saveButton.disabled = pendingSteps.length === 0;
      showStatus(
        pendingSteps.length === 0 ? "error" : "success",
        pendingSteps.length === 0
          ? "Nothing recorded. Run some actions while recording."
          : `Recorded ${pendingSteps.length} steps. Name and save the automation.`
      );
      return;
    }
    recorder.start();
    pendingSteps = null;
    saveButton.disabled = true;
    recordButton.textContent = "Stop";
    recordButton.classList.add("recording");
    showStatus("info", "Recording. Run the actions to capture, then press Stop.");
  });

  saveButton.addEventListener("click", () => {
    void runTask(async () => {
      if (!pendingSteps || pendingSteps.length === 0) throw new Error("Record steps first.");
      const existing = loadAutomations();
      const name = validateName(
        requiredElement<HTMLInputElement>("automation-name").value,
        existing
      );
      saveAutomations([...existing, { name, steps: pendingSteps }]);
      pendingSteps = null;
      saveButton.disabled = true;
      requiredElement<HTMLInputElement>("automation-name").value = "";
      renderAutomations();
      return `Saved "${name}".`;
    });
  });

  renderAutomations();
}

const API_KEY_STORAGE_KEY = "slideware.apiKey";

function loadApiKey(): string {
  try {
    return localStorage.getItem(API_KEY_STORAGE_KEY) ?? "";
  } catch {
    return "";
  }
}

async function selectedTextShape(): Promise<{ id: string; text: string }> {
  const selected = await readSelection();
  const withText = selected.find((shape) => shape.text.trim().length > 0);
  if (!withText) throw new Error("Select a shape that contains text.");
  return { id: withText.id, text: withText.text };
}

function deckOutline(deck: Awaited<ReturnType<typeof snapshotDeck>>): string {
  return deck.slides
    .map((slide) => {
      const texts = slide.shapes
        .map((shape) => shape.text.trim())
        .filter((text) => text.length > 0)
        .join(" | ");
      return `Slide ${slide.index}: ${slide.title ?? "(untitled)"}\n${texts}`;
    })
    .join("\n\n");
}

function bindAiControls(): void {
  const apiKeyInput = requiredElement<HTMLInputElement>("api-key");
  apiKeyInput.value = loadApiKey();
  requiredElement<HTMLButtonElement>("save-api-key").addEventListener("click", () => {
    void runTask(async () => {
      try {
        localStorage.setItem(API_KEY_STORAGE_KEY, apiKeyInput.value.trim());
      } catch {
        throw new Error("This browser blocks storage; the key cannot be saved.");
      }
      return "API key saved to this browser profile.";
    });
  });

  document.querySelectorAll<HTMLButtonElement>("[data-preset]").forEach((button) => {
    button.addEventListener("click", () => {
      void runTask(async () => {
        const target = await selectedTextShape();
        const prompt = presetPrompt(target.text, button.dataset.preset as string);
        const result = await callClaude({
          apiKey: loadApiKey(),
          system: prompt.system,
          messages: [{ role: "user", content: prompt.user }],
        });
        await replaceShapeText(target.id, result);
        return `${button.textContent} applied.`;
      });
    });
  });

  requiredElement<HTMLButtonElement>("run-edit").addEventListener("click", () => {
    void runTask(async () => {
      const instruction = requiredElement<HTMLInputElement>("edit-instruction").value.trim();
      if (!instruction) throw new Error("Describe the edit first.");
      const target = await selectedTextShape();
      const prompt = editPrompt(target.text, instruction);
      const result = await callClaude({
        apiKey: loadApiKey(),
        system: prompt.system,
        messages: [{ role: "user", content: prompt.user }],
      });
      await replaceShapeText(target.id, result);
      return "Edit applied.";
    });
  });

  requiredElement<HTMLButtonElement>("run-create").addEventListener("click", () => {
    void runTask(async () => {
      const topic = requiredElement<HTMLInputElement>("create-topic").value.trim();
      if (!topic) throw new Error("Describe the slide first.");
      const prompt = createPrompt(topic);
      const raw = await callClaude({
        apiKey: loadApiKey(),
        system: prompt.system,
        messages: [{ role: "user", content: prompt.user }],
      });
      const content = parseCreateResponse(raw);
      const brand = loadBrand();
      await insertShapes([
        {
          kind: "textbox",
          left: 60,
          top: 60,
          width: 840,
          height: 50,
          text: content.title,
          fontName: brand.headingFont,
          fontSize: 30,
          fontColor: brand.colors[0],
          bold: true,
        },
        {
          kind: "textbox",
          left: 60,
          top: 140,
          width: 840,
          height: 60 + content.bullets.length * 28,
          text: content.bullets.map((bullet) => `• ${bullet}`).join("\n"),
          fontName: brand.bodyFont,
          fontSize: 18,
          fontColor: brand.colors[0],
        },
      ]);
      return `Created "${content.title}" with ${content.bullets.length} bullets.`;
    });
  });

  requiredElement<HTMLButtonElement>("run-translate").addEventListener("click", () => {
    void runTask(async () => {
      const language = requiredElement<HTMLSelectElement>("translate-language").value;
      const selected = await readSelection();
      const textShapes = selected.filter((shape) => shape.text.trim().length > 0);
      if (textShapes.length === 0) throw new Error("Select shapes that contain text.");
      for (const shape of textShapes) {
        const prompt = translatePrompt(shape.text, language);
        const result = await callClaude({
          apiKey: loadApiKey(),
          system: prompt.system,
          messages: [{ role: "user", content: prompt.user }],
        });
        await replaceShapeText(shape.id, result);
      }
      return `Translated ${textShapes.length} shapes into ${language}.`;
    });
  });
}

function bindSearchControls(): void {
  const run = (): void => {
    void runTask(async () => {
      const query = requiredElement<HTMLInputElement>("search-query").value;
      const deck = await snapshotDeck();
      const hits = searchDeck(deck, query);
      const list = clearList("search-results");
      hits.forEach((hit) => {
        addListItem(list, hit.snippet, `Slide ${hit.slideIndex} · ${hit.shapeName}`, () => {
          void runTask(async () => {
            await gotoSlide(hit.slideIndex);
            if (canSelectShapes()) await setSelection([hit.shapeId]);
            return `Moved to slide ${hit.slideIndex}.`;
          });
        });
      });
      return hits.length === 0 ? "No matches found." : `Found ${hits.length} matches.`;
    });
  };
  requiredElement<HTMLButtonElement>("run-search").addEventListener("click", run);
  requiredElement<HTMLInputElement>("search-query").addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      run();
    }
  });
}

interface ChatTurn {
  role: "user" | "assistant";
  content: string;
}

const darwinHistory: ChatTurn[] = [];

function appendBubble(role: "user" | "assistant", text: string): void {
  const log = requiredElement<HTMLElement>("darwin-log");
  const bubble = document.createElement("div");
  bubble.className = role === "user" ? "bubble bubble-user" : "bubble bubble-assistant";
  bubble.textContent = text;
  log.appendChild(bubble);
  log.scrollTop = log.scrollHeight;
}

function bindDarwinControls(): void {
  const input = requiredElement<HTMLInputElement>("darwin-input");
  const send = (): void => {
    const question = input.value.trim();
    if (!question) return;
    void runTask(async () => {
      appendBubble("user", question);
      input.value = "";
      darwinHistory.push({ role: "user", content: question });
      const deck = await snapshotDeck();
      const answer = await callClaude({
        apiKey: loadApiKey(),
        system: darwinSystem(deckOutline(deck)),
        messages: darwinHistory.map((turn) => ({ role: turn.role, content: turn.content })),
      });
      darwinHistory.push({ role: "assistant", content: answer });
      appendBubble("assistant", answer);
      return "Darwin replied.";
    });
  };
  requiredElement<HTMLButtonElement>("darwin-send").addEventListener("click", send);
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      send();
    }
  });
}

function bindMcpControls(): void {
  const toggle = requiredElement<HTMLButtonElement>("mcp-toggle");
  const statusLabel = requiredElement<HTMLElement>("mcp-status");

  onBridgeStatus((state, detail) => {
    statusLabel.className = state === "connected" ? "mcp-status connected" : "mcp-status";
    statusLabel.textContent =
      state === "connected"
        ? "Connected"
        : state === "connecting"
          ? "Connecting..."
          : "Disconnected";
    toggle.textContent = state === "disconnected" ? "Connect" : "Disconnect";
    if (detail) showStatus("error", detail);
  });

  toggle.addEventListener("click", () => {
    if (isBridgeConnected()) {
      disconnectBridge();
      return;
    }
    connectBridge();
  });
}

function bindTabs(): void {
  const tabs = document.querySelectorAll<HTMLButtonElement>("[data-tab]");
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((other) => other.setAttribute("aria-selected", String(other === tab)));
      document.querySelectorAll<HTMLElement>("[data-panel]").forEach((panel) => {
        panel.hidden = panel.dataset.panel !== tab.dataset.tab;
      });
    });
  });
}

function bindLayoutControls(): void {
  document.querySelectorAll<HTMLButtonElement>("[data-op]").forEach((button) => {
    button.addEventListener("click", () => void runOp(button.dataset.op as string));
  });

  document.querySelectorAll<HTMLButtonElement>("[data-align]").forEach((button) => {
    button.addEventListener("click", () => void runOp(`align.${button.dataset.align}`));
  });

  document.querySelectorAll<HTMLButtonElement>("[data-distribute]").forEach((button) => {
    button.addEventListener("click", () => void runOp(`distribute.${button.dataset.distribute}`));
  });

  requiredElement<HTMLButtonElement>("apply-matrix").addEventListener("click", () => {
    void runOp("matrix", {
      columns: numberValue("matrix-columns"),
      horizontalGap: numberValue("matrix-column-gap"),
      verticalGap: numberValue("matrix-row-gap"),
    });
  });

  requiredElement<HTMLButtonElement>("apply-circle").addEventListener("click", () => {
    void runOp("circle", {
      radius: numberValue("circle-radius"),
      startAngle: numberValue("circle-angle"),
      clockwise: requiredElement<HTMLSelectElement>("circle-direction").value === "clockwise",
    });
  });
}

const shortcutOps: Record<string, string> = {
  l: "align.left",
  c: "align.center",
  r: "align.right",
  t: "align.top",
  m: "align.middle",
  b: "align.bottom",
  h: "distribute.horizontal",
  v: "distribute.vertical",
  w: "size.width",
  e: "size.height",
  s: "swap",
};

function bindShortcuts(): void {
  document.addEventListener("keydown", (event) => {
    const target = event.target as HTMLElement | null;
    const tag = target?.tagName ?? "";
    if (tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA") return;
    if (event.ctrlKey || event.metaKey || event.altKey || busy) return;
    const op = shortcutOps[event.key.toLowerCase()];
    if (!op) return;
    event.preventDefault();
    void runOp(op);
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  requiredElement<HTMLElement>("sideload-msg").hidden = true;
  requiredElement<HTMLElement>("app-body").hidden = false;
  registerLayoutOps();
  registerSelectionOp();
  registerBrandOps();
  registerTemplateOp();
  bindTabs();
  bindLayoutControls();
  bindSelectionControls();
  bindCheckerControls();
  bindBrandControls();
  bindTemplateControls();
  bindAutomationControls();
  bindAiControls();
  bindSearchControls();
  bindDarwinControls();
  bindMcpControls();
  bindShortcuts();
});
