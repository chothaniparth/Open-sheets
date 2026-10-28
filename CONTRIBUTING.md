# Contributing to Open Sheets

Thanks for helping! Bug reports, file samples that don't open correctly, and pull requests are all welcome.

## Getting started

```bash
npm install
npm run dev
```

Before opening a pull request, run:

```bash
npm run typecheck
npm test
```

## Where things live

- **New spreadsheet behaviour** → `src/renderer/src/engine/workbook.ts`. Every change to the document goes through `Workbook.transact()` so it is undoable; HyperFormula changes must happen inside a transaction with `{ hf: true }`.
- **File formats** → `src/renderer/src/io/`. Add a round-trip test in `tests/io.test.ts` for any format change.
- **AI tools** → declare the tool schema in `src/main/ai.ts` and implement it in `src/renderer/src/ai/tools.ts`.
- **AI providers** → `src/main/providers.ts`.

## Good first projects

- Import/export Excel charts as dashboard widgets
- Pivot tables
- Cell comments and notes
- Images and shapes
- Data validation (drop-down lists)
- Named ranges UI

## Legal

Please only contribute code you wrote or that is under a GPL-compatible license. Do not contribute code or assets taken from Microsoft Office or other proprietary software.
