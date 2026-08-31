#!/usr/bin/env node
"use strict";

/*
 * Slideware MCP server.
 * stdio side: Model Context Protocol (JSON-RPC 2.0, one message per line) for Claude.
 * WebSocket side: bridge on ws://127.0.0.1:3711 that the Slideware task pane connects to.
 * tools/call requests are forwarded to the connected pane, which executes them in PowerPoint.
 */

const readline = require("readline");
const { WebSocketServer } = require("ws");
const { handleMessage } = require("./protocol");

const BRIDGE_PORT = Number(process.env.SLIDEWARE_BRIDGE_PORT || 3711);
const CALL_TIMEOUT_MS = 30000;

let paneSocket = null;
let nextCallId = 1;
const pendingCalls = new Map();

function send(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function toolResult(id, text, isError) {
  send({
    jsonrpc: "2.0",
    id,
    result: { content: [{ type: "text", text }], isError: Boolean(isError) },
  });
}

const wss = new WebSocketServer({ host: "127.0.0.1", port: BRIDGE_PORT });

wss.on("connection", (socket) => {
  paneSocket = socket;
  process.stderr.write("Slideware pane connected.\n");

  socket.on("message", (data) => {
    let payload;
    try {
      payload = JSON.parse(data.toString());
    } catch {
      return;
    }
    const pending = pendingCalls.get(payload.callId);
    if (!pending) return;
    pendingCalls.delete(payload.callId);
    clearTimeout(pending.timer);
    toolResult(pending.rpcId, String(payload.result ?? ""), !payload.ok);
  });

  socket.on("close", () => {
    if (paneSocket === socket) paneSocket = null;
    process.stderr.write("Slideware pane disconnected.\n");
  });
});

wss.on("error", (error) => {
  process.stderr.write(`Bridge error: ${error.message}\n`);
});

function forwardToPane(effect) {
  if (!paneSocket || paneSocket.readyState !== 1) {
    toolResult(
      effect.id,
      "No Slideware pane is connected. Open the add-in in PowerPoint and press Connect in the Gen AI tab.",
      true
    );
    return;
  }
  const callId = nextCallId++;
  const timer = setTimeout(() => {
    pendingCalls.delete(callId);
    toolResult(effect.id, "The Slideware pane did not answer within 30 seconds.", true);
  }, CALL_TIMEOUT_MS);
  pendingCalls.set(callId, { rpcId: effect.id, timer });
  paneSocket.send(JSON.stringify({ callId, tool: effect.tool, args: effect.args }));
}

const rl = readline.createInterface({ input: process.stdin, terminal: false });

rl.on("line", (line) => {
  const trimmed = line.trim();
  if (!trimmed) return;
  let message;
  try {
    message = JSON.parse(trimmed);
  } catch {
    send({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "Parse error" } });
    return;
  }
  const { response, effect } = handleMessage(message);
  if (response) send(response);
  if (effect) forwardToPane(effect);
});

rl.on("close", () => {
  wss.close();
  process.exit(0);
});
