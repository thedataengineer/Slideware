export type OpParams = Record<string, unknown>;

export interface OpDefinition {
  label: string;
  recordable?: boolean;
  run: (params?: OpParams) => Promise<string>;
}

export type RecordListener = (id: string, params?: OpParams) => void;

const ops = new Map<string, OpDefinition>();
let recordListener: RecordListener | null = null;

export function registerOp(id: string, definition: OpDefinition): void {
  ops.set(id, definition);
}

export function getOp(id: string): OpDefinition {
  const definition = ops.get(id);
  if (!definition) throw new Error(`Unknown operation: ${id}.`);
  return definition;
}

export function listOps(): { id: string; label: string; recordable: boolean }[] {
  return Array.from(ops.entries()).map(([id, definition]) => ({
    id,
    label: definition.label,
    recordable: definition.recordable === true,
  }));
}

export function setRecordListener(listener: RecordListener | null): void {
  recordListener = listener;
}

export async function dispatch(id: string, params?: OpParams): Promise<string> {
  const definition = getOp(id);
  const message = await definition.run(params);
  if (definition.recordable && recordListener) {
    recordListener(id, params);
  }
  return message;
}
