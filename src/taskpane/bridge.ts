import { audit } from "./features/checker";
import { searchDeck } from "./features/search";
import { dispatch } from "./dispatcher";
import { gotoSlide, readSelection, replaceShapeText, snapshotDeck } from "./powerpoint";

/* global WebSocket */

export type BridgeStatus = "disconnected" | "connecting" | "connected";

interface BridgeCall {
  callId: number;
  tool: string;
  args: Record<string, unknown>;
}

let socket: WebSocket | null = null;
let statusListener: (status: BridgeStatus, detail?: string) => void = () => undefined;

export function onBridgeStatus(listener: (status: BridgeStatus, detail?: string) => void): void {
  statusListener = listener;
}

async function executeTool(tool: string, args: Record<string, unknown>): Promise<string> {
  switch (tool) {
    case "get_deck_outline": {
      const deck = await snapshotDeck();
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
    case "get_selected_shapes":
      return JSON.stringify(await readSelection(), null, 2);
    case "align_shapes":
      return dispatch(`align.${String(args.mode)}`);
    case "distribute_shapes":
      return dispatch(`distribute.${String(args.axis)}`);
    case "arrange_matrix":
      return dispatch("matrix", {
        columns: args.columns,
        horizontalGap: args.horizontalGap ?? 16,
        verticalGap: args.verticalGap ?? 16,
      });
    case "arrange_circle":
      return dispatch("circle", {
        radius: args.radius,
        startAngle: args.startAngle ?? -90,
        clockwise: args.clockwise !== false,
      });
    case "same_size":
      return dispatch(`size.${String(args.mode)}`);
    case "swap_shapes":
      return dispatch("swap");
    case "smart_select":
      return dispatch("select.smart", args);
    case "apply_branding":
      return dispatch(args.scope === "deck" ? "brand.applyDeck" : "brand.applySelection");
    case "insert_template":
      return dispatch("template.insert", { name: args.name });
    case "insert_agenda":
      return dispatch("agenda.insert");
    case "replace_fonts":
      return dispatch("font.replace", { from: args.from, to: args.to });
    case "split_text_box":
      return dispatch("text.split");
    case "merge_text_boxes":
      return dispatch("text.merge");
    case "run_checker": {
      const deck = await snapshotDeck();
      const findings = audit(deck);
      if (findings.length === 0) return "Checker found no issues.";
      return findings
        .map(
          (finding) =>
            `${finding.slideIndex ? `Slide ${finding.slideIndex}: ` : ""}${finding.message}`
        )
        .join("\n");
    }
    case "search_deck": {
      const deck = await snapshotDeck();
      const hits = searchDeck(deck, String(args.query ?? ""));
      if (hits.length === 0) return "No matches found.";
      return hits
        .map((hit) => `Slide ${hit.slideIndex} (${hit.shapeName}): ${hit.snippet}`)
        .join("\n");
    }
    case "goto_slide":
      await gotoSlide(Number(args.index));
      return `Moved to slide ${Number(args.index)}.`;
    case "set_shape_text":
      await replaceShapeText(String(args.shapeId), String(args.text ?? ""));
      return "Text replaced.";
    default:
      throw new Error(`The pane does not implement tool: ${tool}.`);
  }
}

async function handleCall(call: BridgeCall): Promise<void> {
  let ok = true;
  let result: string;
  try {
    result = await executeTool(call.tool, call.args ?? {});
  } catch (error) {
    ok = false;
    result = error instanceof Error ? error.message : "The operation failed in PowerPoint.";
  }
  socket?.send(JSON.stringify({ callId: call.callId, ok, result }));
}

export function connectBridge(url = "ws://127.0.0.1:3711"): void {
  if (socket) return;
  statusListener("connecting");
  const ws = new WebSocket(url);
  socket = ws;

  ws.addEventListener("open", () => statusListener("connected"));
  ws.addEventListener("message", (event) => {
    try {
      const call = JSON.parse(String(event.data)) as BridgeCall;
      void handleCall(call);
    } catch {
      // Ignore malformed frames.
    }
  });
  ws.addEventListener("close", () => {
    if (socket === ws) {
      socket = null;
      statusListener("disconnected");
    }
  });
  ws.addEventListener("error", () => {
    if (socket === ws) {
      socket = null;
      statusListener("disconnected", "Could not reach the MCP server. Is it running?");
    }
  });
}

export function disconnectBridge(): void {
  const ws = socket;
  socket = null;
  ws?.close();
  statusListener("disconnected");
}

export function isBridgeConnected(): boolean {
  return socket !== null && socket.readyState === WebSocket.OPEN;
}
