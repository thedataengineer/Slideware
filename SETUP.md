# Slideware setup

A PowerPoint task pane add-in: align and arrange shapes, apply brand fonts and colors, split and merge text boxes, search and QA the deck, run macros, and rewrite slide text with a local or hosted AI model.

Everything runs from your own machine. There is no hosted build, so each person runs the dev server locally.

## What you need

All three are hard requirements.

| Requirement | Why |
|---|---|
| Node 18 or newer | Build tooling. Developed on Node 22. [ASSUMPTION] 18 is the floor; only 22 has been exercised. |
| PowerPoint desktop on a **Microsoft 365 subscription** | The add-in ships a unified JSON manifest. Perpetual Office 2019, 2021, and 2024 cannot load one at all. |
| PowerPoint API 1.5 or newer | Checked at startup. Below that the task pane refuses to run and says so. |

AI features also need one of: a local Ollama install, an Anthropic API key, or an OpenAI-compatible endpoint. Pick one during setup, not before.

## Run it

```bash
git clone https://github.com/thedataengineer/Slideware.git && cd Slideware && npm install
```

```bash
npm start
```

That generates the manifest, starts the dev server over HTTPS, installs the add-in into your Microsoft 365 account, and launches PowerPoint. Look for a **Slideware** button on the Home tab.

First run asks to trust a local HTTPS certificate. Windows shows a UAC prompt, macOS asks for your keychain password. One time only.

To remove the add-in from your account:

```bash
npm stop
```

## Choose an AI provider

Open the task pane, go to **Gen AI**, then **AI Provider Engine**. Three options:

**Ollama (local, nothing leaves your machine).** Install Ollama, pull a model, leave the URL at its default. The dev server proxies `/ollama` to `http://localhost:11434` for you. This proxy exists because an HTTPS task pane cannot call `http://localhost` directly, so point the field at the default rather than the raw Ollama address.

**Claude.** Paste an Anthropic API key.

**OpenAI-compatible.** Set the base URL and key. Defaults to `https://api.openai.com/v1`.

Keys are stored by the task pane on the machine that entered them and are sent only to the provider you selected.

## Running on a different port

Port 3000 is the default. Override it with an environment variable:

```bash
DEV_SERVER_PORT=4200 npm start
```

`manifest.json` keeps `localhost:3000` as its canonical placeholder and is never edited. The build writes the resolved port into `manifest.local.json`, which is untracked and is what actually gets sideloaded. The webpack dev server reads the same value, so the two cannot drift.

Resolution order: `DEV_SERVER_PORT`, then `config.dev_server_port` in `package.json`, then 3000. An unusable port fails loudly instead of falling back.

## Giving it to someone else

Add them to the private repo, then send them this file. They follow the steps above.

**Change the add-in id first if you share a tenant.** `manifest.json` carries a fixed GUID. Two people installing it in the same Microsoft 365 tenant compete for one identity, and the second install overwrites the first. Different tenants do not collide. Generate a fresh GUID for the `id` field in their copy.

**Their tenant must permit custom app upload.** If policy blocks it, the install fails account-side and no amount of local work fixes it.

## Troubleshooting

**The ribbon shows the wrong buttons or labels.** Do not clear local caches. Unified manifests install into your Microsoft 365 account, not into a local folder, and PowerPoint rebuilds its ribbon from the account copy on every launch. Deleting `Wef/AppCommands`, editing the manifest, or re-copying it into the PowerPoint `wef` folder all get overwritten. Replace the account copy instead:

```bash
npm stop && npm run build && npm start
```

The account rejects an update at a version it already has, so raise `version` in `manifest.json` when the ribbon changes. Give PowerPoint a minute after launch: it rewrites the ribbon cache a little after the window appears, so checking too early looks like a failure.

**"Slideware requires PowerPoint API 1.5."** The host is too old. Update Office, or use a machine on a current Microsoft 365 build.

**"PowerPoint could not read the presentation." or a similar host error.** The message carries the host error code and the API path that failed, for example `(InvalidArgument at Shape.textFrame)`. That path names the failing call and is the place to start.

**The add-in loads but panels do nothing.** The dev server is not running or is on another port. Confirm `https://localhost:<port>/taskpane.html` answers, and that the port matches the one the add-in was installed with.

**Sign-in problems during `npm start`.**

```bash
npm run signin
```

## Driving it from Claude

The **Claude Desktop MCP** panel connects the task pane to an MCP server so Claude can operate the deck. It is optional and off by default; everything else works without it. Setup is in [mcp-server/README.md](mcp-server/README.md).

## Development

```bash
npm test        # unit tests
npm run typecheck
npm run lint
npm run build   # production bundle into dist/
npm run watch   # rebuild on change
```

Pure logic lives in `src/taskpane/features/` and is unit tested. `src/taskpane/powerpoint.ts` is the Office.js boundary and is exercised in PowerPoint, not in tests, so changes there need a real run.
