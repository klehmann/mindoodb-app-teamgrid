# Mindoo TeamGrid

TeamGrid is a collaborative spreadsheet sample app for MindooDB Haven. It stores one workbook in one MindooDB Automerge document and lets multiple users edit cells, rows, columns, worksheet tabs, formulas, and formatting at the same time, on the same device or on different devices, online or offline. When edits sync, Automerge merges them deterministically — formulas keep pointing at the cells the author meant, even if a row or column was inserted or removed concurrently.

The code is intentionally readable rather than feature-complete. It is meant as a tutorial-grade example for platform developers building document-style Haven apps that need granular structured editing.

## Development

```sh
pnpm install
pnpm --filter mindoodb-app-teamgrid dev
pnpm --filter mindoodb-app-teamgrid test
pnpm --filter mindoodb-app-teamgrid build
```

Use `pnpm --filter mindoodb-app-teamgrid dev:local` when you want Vite aliases to point at local MindooDB and App SDK source packages.

TeamGrid is also available as an example preset on the **Applications** page of Haven; one click registers the app, picks a database for it, and opens an empty workbook.

## Haven Capabilities

TeamGrid follows the same host bridge pattern as TeamEdit:

- `read` lists and opens spreadsheet documents.
- `create` enables File / New and File / Import XLSX.
- `update` enables save and grid mutations.
- `delete` enables document deletion.
- `history` enables the revision picker and read-only historical snapshots.
- `attachments` is not used by the first spreadsheet sample.

Host time travel and document revisions are read-only. The app still allows sheet navigation, selection, copy-friendly viewing, formula inspection, and Excel export in those modes.

## Concurrency Model

The hard part of a collaborative spreadsheet is keeping formulas valid when two people change the grid concurrently. Two examples:

1. User A inserts three rows above row 7 on their laptop on a train, while user B reformats column D on their desktop. When their replicas merge, A's `=SUM(D1:D6)` must still refer to the same six cells, even though those cells now have different visible addresses.
2. The same user runs TeamGrid both embedded in the Haven workspace and in a standalone tab. They edit the same workbook in both. Saves must merge cleanly instead of last-writer-wins.

TeamGrid handles both cases the same way:

- **Stable IDs everywhere.** Rows, columns, worksheets, cells, and formula references all use opaque application IDs. The visible `A1` notation is recomputed on every render from the current row and column order — it is a projection of the data, not the storage format. Inserting a row appends a new ID to the order list; existing IDs do not move; existing formulas do not need to be rewritten.
- **Tombstones instead of hard delete.** Rows, columns, and worksheets are marked `deletedAt` rather than physically removed. History stays explainable, and a formula whose target was deleted renders `#REF!` rather than silently retargeting a different cell.
- **Granular JSON patches with `baseHeads`.** Every mutation goes through `useTeamGridDocument.updateGrid`, which records the intent as a `TeamGridOperation`. On save those operations are serialized into a `MindooDBAppJsonPatch` (`set`, `unset`, `listInsert`, `listDelete`) authored against the document version the user actually saw (`baseHeads`). Haven applies and merges the patch against the real Automerge document on its side. Two users typing into different cells produce disjoint paths; two users inserting rows produce list inserts at the same `baseHeads` so the row order interleaves cleanly; a save against an outdated head is rebased rather than rejected. The merged document is what the editor reloads after the save.

This is the same shape TeamEdit uses for collaborative markdown via the SDK's text edit API; TeamGrid is the first sample to drive the structured JSON patch surface.

## Document Model

Each MindooDB document stores one TeamGrid workbook:

- `subject` is the document title shown in file pickers.
- `kind` is `mindoodb.teamgrid`.
- `teamgrid` contains the versioned workbook schema.

Inside `teamgrid.workbook` a workbook is an ordered list of worksheet IDs plus a map of worksheets keyed by ID. Each worksheet contains its own ordered `rowOrder`, `columnOrder`, and maps for row metadata, column metadata, and cells (`cellsById`, keyed by `rowId:columnId`).

## Formulas

The formula subsystem is permissive-license sample code rather than a GPL/commercial engine integration. It currently supports:

- Arithmetic with `+`, `-`, `*`, and `/`.
- Cell references such as `A1`.
- Ranges such as `A1:B4`.
- The functions `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `CONCAT`, and `TODAY`.
- A dependency graph for invalidation.
- Content assist with function signatures and inline help.
- Highlighting referenced cells and ranges while editing formulas.

The parser stores formula source text plus normalized stable-ID references. The function registry in `src/features/formulas/lib/functionRegistry.ts` is the shared source of truth for evaluation and authoring assistance.

## Formatting

Cells separate semantic values from presentation:

- Values can be empty, string, number, or date.
- Number and date values can carry simple display formats — integer, decimal, currency, percent, date, date+time, time.
- Styles include text color, background color, font family, font size, bold, italic, underline, and horizontal and vertical alignment.
- Row heights and column widths are configurable per worksheet.
- Row and column defaults are kept separate from per-cell overrides so broad formatting does not rewrite every cell.

## Excel Interop

TeamGrid imports and exports the .xlsx format through `exceljs`, and it copy/pastes ranges through the OS clipboard in an Excel-friendly way.

- **File import / export.** `File / Import XLSX` builds a TeamGrid workbook envelope from any `.xlsx` file, mapping values, supported formulas, number and date formats, fonts, fills, and alignment. `File / Export XLSX` produces an `.xlsx` blob from the current workbook. See `src/features/xlsx/lib/importWorkbook.ts` and `src/features/xlsx/lib/exportWorkbook.ts`.
- **Clipboard.** Copy and Cut emit three parallel encodings: a TeamGrid-native JSON payload embedded in the HTML clipboard for lossless TeamGrid-to-TeamGrid paste (preserves stable IDs, styles, and formula source); an Excel-compatible HTML body with `x:fmla`, `x:num`, and `x:str` attributes so Excel keeps relative formulas; and a plain TSV fallback for everything else. Paste decodes in the same priority order. See `src/features/clipboard/lib/`.

TeamGrid intentionally supports only a small, well-chosen subset of Excel's functionality — enough to demonstrate the collaboration model on a real-world structured document, not enough to replace Excel.

## Code Layout

The project is organized by feature under `src/features/*`, with `src/app/` for the application shell and `src/shared/` for cross-cutting helpers. The `@/` Vite/TS alias points at `src/`, so feature folders are addressable as `@/features/...` from anywhere.

Inside each feature folder we keep the same three sub-folders so a contributor can find what they need from the feature name alone:

- `composables/` holds Vue composables (`use*.ts`) that own reactive state for that feature. Composables are the only place state should live.
- `components/` holds Single File Components for that feature. Each component imports its composable and renders the corresponding UI; tests live next to the component.
- `lib/` holds plain TypeScript modules: parsers, evaluators, schemas, helpers. Anything that does not depend on Vue belongs here.

`src/app/` only contains the entry point and the root `App.vue`, which is intentionally a thin orchestration layer: it wires composables together, passes their controller objects into feature dialogs, and renders the toolbar/status bar. `src/shared/` is for primitives that genuinely cross feature boundaries (capability checks, theme tokens, default cell dimensions).

```
src/
  app/                          application shell (entry point + root SFC)
  shared/lib/                   cross-feature helpers (capabilities, theme, dimensions)
  features/
    document/                   Haven bridge, save/history, document dialogs
      composables/{useTeamGridDocument,useErrorDialog,useOpenDialog,
                   useDocumentPropertiesDialog,useWorksheetDialogs}.ts
      lib/{teamgridDocument,teamgridOps,viewOpen}.ts
      components/{DocumentRevisionDialog,TagTreeList,ErrorDialog,
                  OpenSpreadsheetDialog,DocumentPropertiesDialog,
                  DeleteSpreadsheetDialog,RenameWorksheetDialog}.vue
    grid/                       spreadsheet surface and formula bar
      composables/{useSelection,useFormulaBarEditing,useFormulaAssistRouter,
                   useGridClipboard,useCellFormatDialog,useInlineCellEditor,
                   useGridSelectionGestures,useColumnRowResize,
                   useGridClipboardBridge}.ts
      components/{GridViewport,FormulaBar,FormulaAssistPanel,
                  WorksheetTabs,CellFormatDialog}.vue
      lib/{gridProjection,cellFormatting}.ts
    formulas/lib/               parser, evaluator, dependency tracking, registry
    clipboard/lib/              TeamGrid / Excel / TSV clipboard payloads
    xlsx/lib/                   .xlsx import and export
  assets/styles/main.css        design tokens, resets, app-shell layout only
                                (component-specific rules live in <style scoped>
                                next to the component they belong to)
```

Pointers into the most important modules:

- `src/features/document/composables/useTeamGridDocument.ts` owns Haven bridge setup, database/document lifecycle, capability gates, time travel, revision snapshots, and `baseHeads`-aware granular saves. Its `TeamGridAppApi` type is the contract every other composable depends on.
- `src/features/document/lib/teamgridDocument.ts` defines the persisted schema, migrations, ID helpers, and default workbook factory.
- `src/features/document/lib/teamgridOps.ts` defines semantic edit operations and serializes them to `MindooDBAppJsonPatch`.
- `src/features/document/lib/viewOpen.ts` builds the dynamic view navigator used by the File / Open dialog.
- `src/features/grid/lib/gridProjection.ts` turns stable IDs into visible rows, columns, and addresses.
- `src/features/grid/lib/cellFormatting.ts` handles value coercion, date/number formatting, and style merging.
- `src/features/formulas/lib/` contains parsing, evaluation, dependency tracking, and function metadata.
- `src/features/clipboard/lib/` contains the serializer/deserializer for the TeamGrid + Excel + TSV clipboard payloads.
- `src/features/xlsx/lib/` contains the `.xlsx` import and export.
- `src/features/grid/components/GridViewport.vue` is the spreadsheet surface; it delegates selection, inline editing, column/row resizing, and native clipboard plumbing to the composables in `src/features/grid/composables/`.
- The dialog SFCs under `src/features/document/components/` and `src/features/grid/components/CellFormatDialog.vue` each take a controller object from their matching composable (`useErrorDialog`, `useOpenDialog`, `useDocumentPropertiesDialog`, `useWorksheetDialogs`, `useCellFormatDialog`) as a prop. The composable owns state and actions; the SFC renders the UI.

## Current Scope

This sample is deliberately small. It does not yet implement virtualization for very large sheets, named ranges in the UI, cross-workbook formulas, conditional formatting, charts, or a full Excel formula language. We are actively exploring more advanced editors — both open source and commercial — to edit Office formats in a data-sovereign way with full concurrency support, but most packages we evaluated are not yet powerful enough for our requirements (in particular: a clean separation between presentation and a stable-ID data model, and being able to drive edits through granular CRDT-friendly operations rather than whole-document replacements). The pieces above can all be layered onto TeamGrid's stable-ID schema and JSON-patch save path without changing the core collaboration model.
