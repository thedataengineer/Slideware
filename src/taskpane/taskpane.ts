import {
  AlignMode,
  DistributionAxis,
  alignShapes,
  arrangeCircle,
  arrangeMatrix,
  distributeShapes,
} from "./alignment";
import { applyLayout } from "./powerpoint";

/* global document, Office, HTMLElement, HTMLInputElement, HTMLButtonElement, HTMLSelectElement */

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
      void execute(`Align ${button.textContent?.toLowerCase()}`, (shapes) =>
        alignShapes(shapes, mode)
      );
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
      })
    );
  });

  requiredElement<HTMLButtonElement>("apply-circle").addEventListener("click", () => {
    void execute("Circle", (shapes) =>
      arrangeCircle(shapes, {
        radius: numberValue("circle-radius"),
        startAngle: numberValue("circle-angle"),
        clockwise: requiredElement<HTMLSelectElement>("circle-direction").value === "clockwise",
      })
    );
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  requiredElement<HTMLElement>("sideload-msg").hidden = true;
  requiredElement<HTMLElement>("app-body").hidden = false;
  bindControls();
});
