export interface AutomationStep {
  op: string;
  params?: Record<string, unknown>;
}

export interface Automation {
  name: string;
  steps: AutomationStep[];
}

export class Recorder {
  private recording = false;
  private steps: AutomationStep[] = [];

  start(): void {
    this.recording = true;
    this.steps = [];
  }

  stop(): AutomationStep[] {
    this.recording = false;
    return this.steps;
  }

  isRecording(): boolean {
    return this.recording;
  }

  recordStep(op: string, params?: Record<string, unknown>): void {
    if (!this.recording) return;
    this.steps.push({ op, params });
  }
}

export function serializeAutomations(automations: Automation[]): string {
  return JSON.stringify(automations);
}

export function parseAutomations(raw: string | null): Automation[] {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(
      (item): item is Automation =>
        typeof item === "object" &&
        item !== null &&
        typeof (item as Automation).name === "string" &&
        Array.isArray((item as Automation).steps)
    );
  } catch {
    return [];
  }
}

export function validateName(name: string, existing: Automation[]): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Name the automation before saving.");
  }
  if (existing.some((automation) => automation.name === trimmed)) {
    throw new Error(`"${trimmed}" already exists.`);
  }
  return trimmed;
}
