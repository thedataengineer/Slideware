# Changelog

## 2.0.0 (2026-09-02)

### Deck from text

Paste prose or an outline and get a whole deck instead of one slide at a time.

The model reads the pasted text, reports what it found, and asks up to four questions about that specific text. Every question arrives with a recommendation already selected, so accepting all of them is a single click. It then drafts a slide plan you reorder, retitle, or trim before anything is created.

Slides are built on layouts from the open deck's own slide master, so a branded template produces branded slides with no restyling. That is the difference from tools that generate into their own renderer and export something PowerPoint-shaped.

New modules, all pure and unit tested: `model-json.ts` recovers JSON from what models actually return, `deck-interview.ts` and `deck-plan.ts` parse the two responses, `deck-prompts.ts` builds both prompts, `slidelayout.ts` matches a layout, `placeholders.ts` picks the title and body shapes. `powerpoint.ts` gains `readLayoutCatalog` and `addSlidesFromPlan`.

### Fixes

- `Shape.textFrame` throws `InvalidArgument` on shapes that do not support one, and an Office.js sync is atomic, so a single picture placeholder blanked an entire deck read. Detail loads now bisect on failure to isolate the shapes the host refuses.
- `Presentation.setSelectedShapes` does not exist in the PowerPoint API. Shape selection is on `Slide`, so Smart Selection and the search jump had never worked. The capability gate also dropped from API 1.6 to 1.5, which is what `Slide.setSelectedShapes` actually requires.
- `deriveTitle` matched "Subtitle" through `includes("title")` and returned the first array hit, so a title slide could report its subtitle as its title.
- Host errors now carry the error code and the failing API path instead of one generic sentence.
- Production builds rewrote only URLs with a trailing slash, so bare ones such as `websiteUrl` kept pointing at localhost.

### Setup and packaging

- The dev server port is configurable through `DEV_SERVER_PORT`. `manifest.json` keeps `localhost:3000` as a placeholder and is never edited; the resolved port is written into an untracked `manifest.local.json`.
- Each person who installs gets their own add-in id, generated on first run and kept in an untracked `.addin-id`, so a team can sideload from one tenant without sharing an identity.
- The unused `commands` entry point from the project template is gone, along with the ribbon button that pointed at it.
- The ribbon button, manifest metadata, and icons now identify the add-in as Slideware.
- Added README, SETUP, LICENSE (MIT), and NOTICE. The repository is public.

### Known limits

- Slides are appended to the end of the deck. `slides.add` cannot insert, and `Slide.moveTo` needs API 1.8, above the 1.5 floor.
- Speaker notes are written by the model and shown in the plan, but PowerPoint's add-in API has no notes support at any version. They are stored as slide tags: lossless in the file, invisible in the notes pane.
- The deck-building path runs against PowerPoint and is outside the unit test config, so it is verified by running the add-in rather than by `npm test`.
