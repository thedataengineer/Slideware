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
import { dispatch, registerOp } from "./dispatcher";
import {
  applyLayout,
  canSelectShapes,
  gotoSlide,
  readSelection,
  setSelection,
  snapshotDeck,
} from "./powerpoint";

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
  const classes = { info: "status", success: "status status-success", error: "status status-error" };
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
        const meta = finding.slideIndex ? `Slide ${finding.slideIndex} · ${finding.rule}` : finding.rule;
        const jump = finding.slideIndex ? () => void runTask(async () => {
          await gotoSlide(finding.slideIndex as number);
          return `Moved to slide ${finding.slideIndex}.`;
        }) : undefined;
        addListItem(list, finding.message, meta, jump);
      });
      return findings.length === 0
        ? "Checker found no issues."
        : `Checker found ${findings.length} issues.`;
    });
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
  bindTabs();
  bindLayoutControls();
  bindSelectionControls();
  bindCheckerControls();
  bindShortcuts();
});
