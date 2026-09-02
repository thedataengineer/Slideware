# Slideware

A PowerPoint task pane add-in for people who build decks all day. Paste a briefing and get a deck built on your own template. Align and arrange shapes, apply brand fonts and colors, split and merge text boxes, search and QA the deck, record macros, and rewrite slide text with a local or hosted AI model.

Runs entirely from your own machine. There is no hosted service, no telemetry, and no account to create.

## Disclaimer

**This software is provided as is, with no warranty of any kind and no liability.** See [LICENSE](LICENSE) for the full text.

It is a personal project, not a product. It is not affiliated with, endorsed by, or supported by Microsoft or Anthropic. It edits your presentations directly through the Office JavaScript API, including moving, resizing, deleting, and rewriting shapes. Several of those operations cannot be undone from inside the add-in. **Work on a copy of anything you cannot afford to lose.**

If you enable an AI provider, slide text is sent to whichever provider you configure. Choose Ollama to keep that on your own machine.

## What it does

**Productivity.** Align and distribute selected shapes. Arrange them into a grid or a radial circle. Split one text box into lines or merge several in order. Fit text to a box or a box to its text. Select every shape matching an anchor by type, size, fill, font, font color, or font size. Search the whole deck and jump to a hit. Run a QA pass and insert an agenda. Insert slide templates. Record a sequence of actions and replay it as a macro. Single-key shortcuts for the common alignment moves.

**Branding.** Apply brand fonts and colors across a selection or the whole deck. Replace one font with another everywhere.

**Deck from text.** Paste prose or an outline and get a whole deck. The model reads the text, reports what it found, and asks up to four questions about that specific text with a recommendation already selected, so accepting all of them is one click. It then drafts a slide plan you reorder, retitle, or trim before anything is created. Slides are built on layouts from your deck's own slide master, so a branded template comes out branded with no restyling.

**Gen AI.** Rewrite, shorten, or edit slide text with Claude, an OpenAI-compatible endpoint, or a local Ollama model. Generate a single slide from a topic, grounded in the deck outline. Translate a selection. Ask a deck assistant about the open presentation.

**MCP.** Optionally lets Claude drive the add-in directly. See [mcp-server/README.md](mcp-server/README.md).

## Requirements

- Node 18 or newer
- PowerPoint desktop on a **Microsoft 365 subscription**. Perpetual Office 2019, 2021, and 2024 cannot load a unified manifest.
- PowerPoint API 1.5 or newer

## Getting started

```bash
git clone https://github.com/thedataengineer/Slideware.git && cd Slideware && npm install && npm start
```

Full instructions, provider setup, port configuration, sharing with a team, and troubleshooting are in [SETUP.md](SETUP.md).

Installing the add-in uploads it to your Microsoft 365 account. Many corporate tenants disable custom app upload, which blocks the install account-side. Check that before planning a team rollout.

## Development

```bash
npm test        # unit tests
npm run typecheck
npm run lint
npm run build   # production bundle into dist/
```

Pure logic lives in `src/taskpane/features/` and is unit tested. `src/taskpane/powerpoint.ts` is the Office.js boundary; it is exercised in PowerPoint rather than in tests, so changes there need a real run before you trust them.

## Changes

See [CHANGELOG.md](CHANGELOG.md).

## License

MIT. See [LICENSE](LICENSE). Third party attribution is in [NOTICE](NOTICE).
