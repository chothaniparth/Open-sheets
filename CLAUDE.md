# CLAUDE.md

Open Sheets is an open-source macOS spreadsheet app built with Electron, React and HyperFormula. It is a clean-room build, not derived from Excel. It is licensed GPL-3.0-or-later, which is required because HyperFormula runs under its `gpl-v3` license key.

## Commands

```bash
npm run dev          # electron-vite dev with HMR
npm run build        # bundle main/preload/renderer into out/
npm start            # run the built app (electron-vite preview)
npm run typecheck    # tsc --noEmit (TypeScript 7 / tsgo)
npm test             # vitest: engine, io round-trips, AI tools
npm run dist:mac:unsigned   # .app + DMG/zip without a signing identity
node scripts/make-sample.mjs # regenerate samples/Sales Demo.xlsx
```

- **Launching Electron from VS Code:** the integrated terminal sets `ELECTRON_RUN_AS_NODE=1`, which breaks Electron imports with "does not provide an export named BrowserWindow". The npm scripts already run with `env -u ELECTRON_RUN_AS_NODE`. Do the same when you launch Electron by hand.
- **Electron binary download:** if `npm install` cannot fetch the Electron binary, download `electron-v45.0.0-darwin-arm64.zip` with curl. Unpack it with `ditto` into `node_modules/electron/dist`, then write `Electron.app/Contents/MacOS/Electron` to `node_modules/electron/path.txt`.
- **Building DMGs locally:** electron-builder's `dmg` target can fail while downloading its helpers. Instead, build the app bundle with `npx electron-builder --mac dir --arm64 -c.mac.identity=null -c.electronDist=node_modules/electron/dist`. Then make the DMG with `hdiutil create -format UDZO` from a folder containing the `.app` and an `/Applications` symlink. For x64, pass `--x64` and point `-c.electronDist` at an unpacked darwin-x64 Electron.

## Debugging the real app

`src/main/index.ts` reads these environment variables:

- `OPEN_SHEETS_DEBUG=1` forwards renderer console output to stdout.
- `OPEN_SHEETS_CAPTURE=/path/out.png` writes a screenshot of the window and quits. `OPEN_SHEETS_CAPTURE_DELAY` sets the wait in ms (default 2500). Allow about 15000 or more for the x64 build under Rosetta.
- `OPEN_SHEETS_CAPTURE_JS='...'` runs in the renderer before the capture, for example to click a tab.

Pass a file path as an argument to open it, e.g. `env -u ELECTRON_RUN_AS_NODE npx electron . "samples/Sales Demo.xlsx"`.

## Architecture

### Main process (`src/main`, ESM)

- `index.ts`: windows, file open/save IPC, close-with-unsaved-changes prompt, native context menus, and printing and PDF export through a hidden window. Also enforces a single instance and handles `open-file`.
- `menu.ts`: the macOS menu bar. Each item sends a `menu:command` id, and `run(cmd)` in `App.tsx` handles it.
- `store.ts`:
  - `settings.json` holds the settings.
  - `secrets.json` holds API keys encrypted with `safeStorage` (Keychain).
  - Keys never reach the renderer. The renderer only receives `hasKey` and `keyHint`.
- `providers.ts`: the built-in provider list and the `createModel(provider, modelId)` factory, built on the Vercel AI SDK v7 `@ai-sdk/*` packages. `fetchModelList` fetches each provider's model list.
- `ai.ts`:
  - The chat loop is `streamText` with zod tools.
  - Tool `execute` functions do not touch data. They send `ai:tool-request` to the renderer and wait up to 60s for `ai:tool-response`. The workbook lives only in the renderer.
  - Sessions are keyed `${webContents.id}:${chatId}`.
  - **To add an AI tool:** define it in `ai.ts` and implement it in `src/renderer/src/ai/tools.ts` (`runTool`).

### Preload (`src/preload/index.ts`)

The preload is built as **CJS** (`index.cjs`) because the window is sandboxed. It exposes `window.api` (`file`, `win`, `app`, `settings`, `ai`), and the types are in `src/renderer/src/env.d.ts`.

### Shared (`src/shared/types.ts`)

Holds the IPC payload types and `FILE_EXTENSIONS`, the open and save lists. Import it via the `@shared/*` alias.

### Renderer (`src/renderer/src`, React 19)

#### `engine/`: framework-free and unit-tested

`workbook.ts` wraps HyperFormula and stores per-sheet metadata: styles, sizes, merges, freeze panes, filters, conditional formatting, plus dashboards. Rules:

- **Every mutation goes through `wb.transact(label, fn)`.** It snapshots the metadata and pairs it with HyperFormula's undo entry, so undo and redo stay in sync.
- `hf.batch()` always adds exactly one undo entry.
- You cannot read values inside a batch.
- Row and column insert/delete must shift every piece of metadata, as the existing methods do.

The other engine files:

- `addr.ts`: A1 addressing, ranges, and `shiftFormula` for copy, paste and fill (respects `$` anchors; out-of-grid references become `#REF!`).
- `format.ts`: number formats (SSF), input parsing, and date serials (epoch 1899-12-30).
- `axis.ts`: row and column sizes for the virtualized grid.

#### `io/`

- `importer.ts` reads xlsx and its macro/template variants with ExcelJS, keeping styles. It uses SheetJS for everything else, and as a fallback when ExcelJS fails.
- `exporter.ts` writes xlsx via ExcelJS and the other formats via SheetJS. `sheetToHtml` also renders print output.
- Dashboards, conditional formatting and tab colors round-trip through the very-hidden sheet `_opensheets_meta`, stored as JSON chunked across cells.

#### Other renderer parts

- `components/Grid.tsx`: the canvas-rendered grid, with 4 freeze regions and a scroller overlay. It handles selection, editing, reference picking while typing formulas, the fill handle, and resizing.
- `charts/options.ts`: builds ECharts options. Keep the existing categorical palettes (`SERIES_LIGHT` / `SERIES_DARK`).
- `components/Dashboard.tsx`: react-grid-layout **v2** API (`GridLayout`, `useContainerWidth`, `gridConfig`/`dragConfig`).
- `App.tsx`: the central component. It owns state, the `run(cmd)` command switch, clipboard handling, saving, and the bridge from AI tool calls to the renderer (`hostRef`).
- `styles.css`: light and dark tokens, switched with `data-theme` on `<html>`.

## Conventions

- Match the existing style: no semicolons, single quotes, 2-space indent, sparse comments.
- Icons come from `lucide-react`. Check that an icon name exists before using it, because several were renamed (e.g. `TextAlignStart`, `Funnel`, `Trash`).
- Renderer-only libraries are `devDependencies` because they get bundled. Main-process runtime dependencies (`ai`, `@ai-sdk/*`, `zod`) are `dependencies`.
- SheetJS is installed from the SheetJS CDN tarball (0.20.3), not from npm.
- Claude models: the defaults are `claude-opus-5-5`, `claude-sonnet-5-5` and `claude-haiku-5-5`. Opus 5.5 rejects `temperature` and forced `tool_choice`, so don't send them.
- Run `npm run typecheck && npm test` before finishing a change. To confirm a UI change, build and capture a screenshot of the real app with `OPEN_SHEETS_CAPTURE`.

## Release

- `.github/workflows/release.yml` runs on `v*` tags and builds the arm64 and x64 DMGs and zips on macOS.
- It signs and notarizes when the `MAC_CERT_P12_BASE64` and Apple ID secrets are set. Otherwise it builds unsigned.
- `website.yml` deploys `website/` to GitHub Pages.
