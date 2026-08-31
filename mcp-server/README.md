# Slideware MCP server

Lets Claude (Claude Desktop, Claude Code, or any MCP client) drive the Slideware add-in inside PowerPoint: read the deck outline, align and arrange shapes, apply branding, insert templates, run the checker, search slides, and edit shape text.

## How it works

```
Claude ──stdio (MCP)──> mcp-server/server.js ──ws://127.0.0.1:3711──> Slideware task pane ──Office.js──> PowerPoint
```

The server speaks MCP over stdio. Tool calls are forwarded over a local WebSocket to the Slideware task pane, which executes them with Office.js and returns the result. Both halves run on your machine; nothing leaves it.

## Setup

1. Install dependencies once:

   ```bash
   cd mcp-server && npm install
   ```

2. Register the server with Claude.

   **Claude Code:**

   ```bash
   claude mcp add slideware -- node /absolute/path/to/Slideware/mcp-server/server.js
   ```

   **Claude Desktop** (`claude_desktop_config.json`):

   ```json
   {
     "mcpServers": {
       "slideware": {
         "command": "node",
         "args": ["/absolute/path/to/Slideware/mcp-server/server.js"]
       }
     }
   }
   ```

3. Open the Slideware add-in in PowerPoint, go to the **Gen AI** tab, and press **Connect** under Claude MCP. The status dot turns green when the pane is linked.

4. Ask Claude things like "align the selected shapes left", "run the deck checker", or "search the deck for revenue".

## Notes

- Tool calls fail with a clear message when no pane is connected or the pane does not answer within 30 seconds.
- Set `SLIDEWARE_BRIDGE_PORT` to change the WebSocket port (default 3711); set the same port in the task pane connect field if you do.
