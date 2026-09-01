"use strict";

const PROTOCOL_VERSION = "2024-11-05";

const TOOLS = [
  {
    name: "get_deck_outline",
    description: "Read every slide's title and text as an outline of the open presentation.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "get_selected_shapes",
    description: "Read the currently selected shapes (position, size, text, formatting).",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "align_shapes",
    description: "Align the selected shapes to a selection edge or centerline.",
    inputSchema: {
      type: "object",
      properties: { mode: { type: "string", enum: ["left", "center", "right", "top", "middle", "bottom"] } },
      required: ["mode"],
      additionalProperties: false,
    },
  },
  {
    name: "distribute_shapes",
    description: "Distribute the selected shapes with equal gaps along an axis (needs 3+ shapes).",
    inputSchema: {
      type: "object",
      properties: { axis: { type: "string", enum: ["horizontal", "vertical"] } },
      required: ["axis"],
      additionalProperties: false,
    },
  },
  {
    name: "arrange_matrix",
    description: "Arrange the selected shapes into a grid with the given columns and gaps.",
    inputSchema: {
      type: "object",
      properties: {
        columns: { type: "integer", minimum: 1 },
        horizontalGap: { type: "number" },
        verticalGap: { type: "number" },
      },
      required: ["columns"],
      additionalProperties: false,
    },
  },
  {
    name: "arrange_circle",
    description: "Arrange the selected shapes evenly around a circle.",
    inputSchema: {
      type: "object",
      properties: {
        radius: { type: "number", exclusiveMinimum: 0 },
        startAngle: { type: "number" },
        clockwise: { type: "boolean" },
      },
      required: ["radius"],
      additionalProperties: false,
    },
  },
  {
    name: "same_size",
    description: "Match the selected shapes' width, height, or both to the first selected shape.",
    inputSchema: {
      type: "object",
      properties: { mode: { type: "string", enum: ["width", "height", "both"] } },
      required: ["mode"],
      additionalProperties: false,
    },
  },
  {
    name: "swap_shapes",
    description: "Swap the positions of exactly two selected shapes.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "smart_select",
    description: "Select all shapes on the anchor shape's slide matching the given criteria.",
    inputSchema: {
      type: "object",
      properties: {
        sameType: { type: "boolean" },
        sameFill: { type: "boolean" },
        sameSize: { type: "boolean" },
      },
      additionalProperties: false,
    },
  },
  {
    name: "apply_branding",
    description: "Apply the saved brand: fonts and colors to the selection, or fonts across the deck.",
    inputSchema: {
      type: "object",
      properties: { scope: { type: "string", enum: ["selection", "deck"] } },
      required: ["scope"],
      additionalProperties: false,
    },
  },
  {
    name: "insert_template",
    description: "Insert a brand-styled template block on the current slide.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", enum: ["title-block", "kpi-row", "quote-card", "section-divider"] },
      },
      required: ["name"],
      additionalProperties: false,
    },
  },
  {
    name: "insert_agenda",
    description: "Build an agenda from slide titles and insert it on the current slide.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "run_checker",
    description: "Audit the deck: off-slide shapes, tiny fonts, font sprawl, empty text, overlong text.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "search_deck",
    description: "Search all slide text and return matches with slide numbers and snippets.",
    inputSchema: {
      type: "object",
      properties: { query: { type: "string" } },
      required: ["query"],
      additionalProperties: false,
    },
  },
  {
    name: "goto_slide",
    description: "Move the PowerPoint view to the given 1-based slide number.",
    inputSchema: {
      type: "object",
      properties: { index: { type: "integer", minimum: 1 } },
      required: ["index"],
      additionalProperties: false,
    },
  },
  {
    name: "split_text_box",
    description: "Split the selected multi-line text box into one text box per line.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "merge_text_boxes",
    description: "Merge the selected text boxes into one, joining text in selection order.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "set_shape_text",
    description: "Replace the text of a shape by its id (ids come from other tools).",
    inputSchema: {
      type: "object",
      properties: { shapeId: { type: "string" }, text: { type: "string" } },
      required: ["shapeId", "text"],
      additionalProperties: false,
    },
  },
];

const TOOL_NAMES = new Set(TOOLS.map((tool) => tool.name));

function response(id, result) {
  return { jsonrpc: "2.0", id, result };
}

function errorResponse(id, code, message) {
  return { jsonrpc: "2.0", id, error: { code, message } };
}

function handleMessage(message) {
  const { id, method, params } = message;

  if (method === "initialize") {
    return {
      response: response(id, {
        protocolVersion: PROTOCOL_VERSION,
        capabilities: { tools: {} },
        serverInfo: { name: "slideware", version: "1.0.0" },
      }),
    };
  }

  if (typeof method === "string" && method.startsWith("notifications/")) {
    return {};
  }

  if (method === "ping") {
    return { response: response(id, {}) };
  }

  if (method === "tools/list") {
    return { response: response(id, { tools: TOOLS }) };
  }

  if (method === "tools/call") {
    const name = params && params.name;
    if (!TOOL_NAMES.has(name)) {
      return { response: errorResponse(id, -32602, `Unknown tool: ${name}`) };
    }
    return { effect: { type: "forward", id, tool: name, args: (params && params.arguments) || {} } };
  }

  return { response: errorResponse(id, -32601, `Method not found: ${method}`) };
}

module.exports = { PROTOCOL_VERSION, TOOLS, handleMessage };
