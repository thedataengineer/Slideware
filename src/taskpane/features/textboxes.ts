import { SnapshotShape } from "./snapshot";

export interface TextBoxSpec {
  kind: "textbox";
  left: number;
  top: number;
  width: number;
  height: number;
  text: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
}

export interface SplitResult {
  deleteIds: string[];
  inserts: TextBoxSpec[];
}

export interface MergeResult {
  deleteIds: string[];
  insert: TextBoxSpec;
}

function formatOf(shape: SnapshotShape): Pick<TextBoxSpec, "fontName" | "fontSize" | "fontColor"> {
  return { fontName: shape.fontName, fontSize: shape.fontSize, fontColor: shape.fontColor };
}

export function splitTextBox(shape: SnapshotShape): SplitResult {
  const lines = shape.text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (lines.length < 2) {
    throw new Error("Select a text box with more than one line to split.");
  }

  const lineHeight = shape.height / lines.length;
  return {
    deleteIds: [shape.id],
    inserts: lines.map((line, index) => ({
      kind: "textbox",
      left: shape.left,
      top: shape.top + index * lineHeight,
      width: shape.width,
      height: lineHeight,
      text: line,
      ...formatOf(shape),
    })),
  };
}

export function mergeTextBoxes(shapes: SnapshotShape[]): MergeResult {
  const withText = shapes.filter((shape) => shape.text.trim().length > 0);
  if (withText.length < 2) {
    throw new Error("Select at least 2 text boxes with text to merge.");
  }

  const left = Math.min(...withText.map((shape) => shape.left));
  const top = Math.min(...withText.map((shape) => shape.top));
  const right = Math.max(...withText.map((shape) => shape.left + shape.width));
  const bottom = Math.max(...withText.map((shape) => shape.top + shape.height));

  return {
    deleteIds: withText.map((shape) => shape.id),
    insert: {
      kind: "textbox",
      left,
      top,
      width: right - left,
      height: bottom - top,
      text: withText.map((shape) => shape.text.trim()).join("\n"),
      ...formatOf(withText[0]),
    },
  };
}
