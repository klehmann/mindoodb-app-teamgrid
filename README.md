# Mindoo TeamGrid

TeamGrid is a collaborative spreadsheet sample app for MindooDB Haven. It stores one workbook in one MindooDB Automerge document and lets multiple users edit cells, rows, columns, worksheet tabs, formulas, formatting, and embedded charts at the same time, on the same device or on different devices, online or offline. When edits sync, Automerge merges them deterministically — formulas keep pointing at the cells the author meant, even if a row or column was inserted or removed concurrently.

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
- `create` enables File / New, File / New from template..., and File / Import XLSX.
- `update` enables save and grid mutations.
- `delete` enables document deletion.
- `history` enables the revision picker and read-only historical snapshots.
- `views` enables the File / Open dialog's view-backed category tree (driven by the navigator API) and Virtual View Sheets (see below) that materialize MindooDB virtual view data into worksheet rows.
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
- `tags` is a normalized list of hierarchical tag strings (`Work\Planning`) used by the File / Open category tree.
- `istemplate` marks the spreadsheet as a template so it can be filtered in File / Open and duplicated via File / New from template... (see [Spreadsheet Templates](#spreadsheet-templates)).
- `kind` is `mindoodb.teamgrid`.
- `teamgrid` contains the versioned workbook schema.

Inside `teamgrid.workbook` a workbook is an ordered list of worksheet IDs plus a map of worksheets keyed by ID. Each worksheet contains its own ordered `rowOrder`, `columnOrder`, and maps for row metadata, column metadata, and cells (`cellsById`, keyed by `rowId:columnId`). A worksheet may also carry an optional `viewBinding` that turns it into a Virtual View Sheet — see [Virtual View Sheets](#virtual-view-sheets).

## Formulas

TeamGrid uses the permissively licensed `fast-formula-parser` package behind a local adapter, so the app now supports a much broader Excel-compatible formula subset while keeping TeamGrid's stable-ID storage model. Supported formulas include arithmetic, comparisons, cell references, ranges, whole-column references, formula-to-formula dependencies, and hundreds of functions from categories such as math, statistics, text, date/time, and logical formulas. Examples include `IF`, `ROUND`, `COUNTIF`, `LOWER`, `UPPER`, `LEFT`, `RIGHT`, `DATE`, `YEAR`, `MONTH`, and `DAY`, in addition to the original `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `CONCAT`, and `TODAY`.

The adapter in `src/features/formulas/lib/fastFormulaParserAdapter.ts` maps visible Excel-style addresses (`A1`, `A1:B4`, `A:A`) to TeamGrid's persisted row and column IDs. Formula cells store the original source text plus normalized stable-ID references, so formulas keep pointing at the intended cells when rows or columns are inserted, deleted, or concurrently reordered.

Formulas can also reference cells and ranges on other worksheets. Use Excel's sheet qualifier syntax, quoting sheet names that contain spaces or punctuation: `=SUM('Sheet 2'!A1:C1)`. TeamGrid accepts the friendlier unquoted form while editing when the sheet name is unambiguous, such as `=SUM(Sheet 2!A1:C1)`, and normalizes it back to Excel-compatible quoted syntax for display and XLSX export. Internally those cross-sheet references are still stored as worksheet, row, and column IDs, so they tolerate sheet renames and row/column insertions in the target sheet.

Formula authoring now includes Ctrl+Space / fx content assist with live filtering, scrollable suggestions, hover feedback, signatures, and short descriptions for available functions. The metadata in `src/features/formulas/lib/functionRegistry.ts` drives that help UI; evaluation itself is delegated to `fast-formula-parser` through the adapter.

## Formatting

Cells separate semantic values from presentation:

- Values can be empty, string, number, or date.
- Number and date values can carry simple display formats — integer, decimal, currency, percent, date, date+time, time — and an optional Excel `numFmt` string for round-trip fidelity with imported workbooks.
- Styles include text color, background color, font family, font size, bold, italic, underline, and horizontal and vertical alignment. The default font is Calibri 11pt to match Excel; family names are stored bare (`"Calibri"`, `"Arial"`, …) the way Excel does, and the render layer appends a generic CSS fallback at paint time so the document model stays portable. The five Microsoft Office / Windows families that show up most often in XLSX workbooks — Calibri, Cambria, Courier New, Arial, and Times New Roman — are aliased via `@font-face` rules in `main.css` to Google's bundled metric-compatible Croscore replacements (Carlito, Caladea, Cousine, Arimo, Tinos), so column widths and line heights stay stable on macOS, Linux, Windows, and mobile even when the real Office font is not installed.
- Per-edge cell borders carry their own style (`thin`, `medium`, `thick`, `dashed`, `dotted`, `double`) and color. The Format cells dialog exposes range-level presets — none, outline, inside, all — alongside an interactive line picker that toggles individual top / bottom / left / right / inside-horizontal / inside-vertical edges.
- Row heights and column widths are configurable per worksheet.
- Row and column defaults are kept separate from per-cell overrides so broad formatting does not rewrite every cell.

The Format cells dialog (`CellFormatDialog.vue`, controlled by `useCellFormatDialog`) is the single entry point for value formatting and presentation:

- **Cell type** picks the value kind (text, general, integer, decimal, percent, currency, date, date+time, time, custom) and accepts a custom Excel `numFmt` string for advanced cases.
- **Alignment** picks horizontal alignment (General — Excel's value-aware default that right-aligns numbers and dates and left-aligns text — plus explicit Left, Center, Right), vertical alignment (Top, Middle, Bottom), an Excel-style indent (0..15, applies to Left and Right), and a Wrap text checkbox that lets multi-line content (Alt+Enter) and long values wrap inside the cell instead of spilling into the empty neighbour.
- **Font** sets family, size, weight, slant, underline, and text color.
- **Fill** toggles and picks the background color.
- **Border** picks edge style and color, applies presets, and edits individual edges.

Every action in the dialog is range-aware: it iterates over the active range plus any additional ranges from multi-selection (see below), so a single Apply mutates every selected rectangle in one granular operation rather than one cell at a time.

## Charts

TeamGrid supports the four chart types most spreadsheets use day to day:

- **Column** charts (vertical bars)
- **Bar** charts (horizontal bars)
- **Line** charts
- **Pie** charts

Charts are first-class objects on each worksheet, stored beside `cellsById` in the stable-ID schema (`chartsById`, `chartOrder`). Each chart has a two-cell anchor on the sheet, a type, optional title, legend settings, category-axis range, and one or more series with name, values range, and optional color.

- **Creating charts.** Select a data range and use `Insert` → Column / Bar / Line / Pie chart. TeamGrid infers category labels and value series from the selection layout (header row/column when present).
- **Editing charts.** Double-click a chart or open Chart properties to change the title, chart type, chart data range, per-series name and values, legend position, series colors, palette colors, and value-axis gridlines.
- **Rendering.** `ChartOverlay.vue` draws charts on top of the grid with D3 scales; bar, column, and line charts show a series legend (matching Excel’s right-side default); pie charts label slices from the category range.
- **Excel round-trip.** `.xlsx` import reads OOXML chart and drawing parts into the stable-ID model; export writes charts back into the workbook’s drawing XML. Optional fields without a value (for example a series with no custom fill color) are omitted rather than stored as `undefined`, because Automerge rejects explicit `undefined` assignments.

See `src/features/charts/` for overlay rendering, data resolution, geometry, and `src/features/xlsx/lib/chartImporter.ts` / `chartExporter.ts` for Excel interop.

## Virtual View Sheets

A Virtual View Sheet is a regular TeamGrid worksheet that is populated from a MindooDB **virtual view** configured for the TeamGrid application in Haven. The host advertises configured views through the launch context; the app turns the user's choice into an opened view navigator on the Haven side, paginates through every entry, and materializes the result into the worksheet's stable-ID row, column, and cell maps. Once materialized, the data is a normal part of the workbook: Excel formulas can reference it, charts can be drawn over it, and the cells participate in the same selection, formatting, copy/paste, undo, and granular-save plumbing as hand-entered data.

This is the bridge that lets a TeamGrid workbook process MindooDB document data with the full spreadsheet feature set — including Excel-compatible formulas, embedded charts, and XLSX export — without giving up the indexed virtual view performance on the database side.

- **Adding a sheet.** The "+" tab menu offers _Add Virtual View Sheet_ alongside _Add Sheet_. The dialog (`ViewSheetSettingsDialog.vue`) picks one of the configured virtual views, a sheet name, whether to show documents, categories, or both, and an optional top-level category path (`somecategory\sublevel1`) that scopes the navigator. The same dialog opens from the tab context menu to edit an existing view sheet.
- **Materialization.** `src/features/view-sheets/lib/viewSheetMaterialization.ts` walks the view navigator with `category_then_document` ordering, projects each navigator entry into a worksheet row, and reads column values straight from `MindooDBAppViewEntry.columnValues`. View-language count formulas (`v.childCount()`, `v.descendantDocumentCount()`, `v.siblingCount()`, …) are resolved against the per-entry counts the navigator already publishes, so categories show pre-aggregated totals without re-counting client-side.
- **Permanent, static cells.** The materialization writes plain `Cell` records into the worksheet just like any other content — there is no special "live view" cell kind. The values are persisted into the MindooDB document, so the data is available offline, participates in revisions and time travel, and stays static between refreshes. This is what lets standard Excel formulas like `=SUM(A2:A20)` and inserted charts reference virtual view rows by ordinary range syntax.
- **Refresh.** Each view sheet exposes a Refresh button (and a "stale" indicator) in the worksheet header. Refresh re-runs the navigator, rewrites the row, column, and cell maps in place, and records the refresh timestamp on the worksheet's `viewBinding`. Manually edited cells in a view sheet are intentionally overwritten on refresh — the source of truth is the configured virtual view.
- **XLSX export.** Because view sheet rows are materialized into the normal cell map, `File / Export XLSX` writes them out as static cell values with no special handling. Recipients open the workbook in Excel and see exactly the same values, formats, and charts the author saw at the moment of the last refresh, with no MindooDB connection required.

See `src/features/view-sheets/` for materialization, and `viewBinding` on `Worksheet` for the persisted reference back to the source view.

## Spreadsheet Templates

Any TeamGrid spreadsheet can be flagged as a template by toggling _Use this spreadsheet as a template_ in the Spreadsheet Properties dialog. The flag is stored as a top-level `istemplate` field on the document envelope (next to `subject` and `tags`) and is persisted through the same `setDocumentProperties` granular operation as the title and tags, so concurrent edits to the flag merge cleanly with other property edits.

- **File / Open filtering.** The Open dialog has a _Spreadsheet type_ selector with three modes — _All_, _No templates_, and _Only templates_. Each mode is backed by a fixed-id dynamic virtual view definition with a different filter expression (`v.neq(v.field("istemplate"), true)` and `v.eq(v.field("istemplate"), true)`), so the category tree, document count badges, and tag fan-out all stay correct without the client filtering documents after the fact. By default the dialog opens in _No templates_ mode so templates do not clutter the normal Open list.
- **File / New from template....** The File menu adds a _New from template..._ command that reuses the Open dialog in a template-picking mode: the type selector is locked to _Only templates_, the dialog header changes to _New from template_, and the confirm button becomes _Create_. Picking a template fetches the source document, clones its Teamgrid envelope, clears `istemplate` on the copy, and titles the new spreadsheet `Copy of <template title>` before loading it in the editor as an ordinary new spreadsheet that the user is free to save back to any writable database.

The template logic lives in `src/features/document/lib/teamgridDocument.ts` (model defaults and `readIsTemplate`), `src/features/document/lib/teamgridOps.ts` (the `setDocumentProperties` payload), `src/features/document/lib/viewOpen.ts` (the three filter-mode view definitions), `src/features/document/composables/useOpenDialog.ts` (open vs. template mode), and `src/features/document/composables/useTeamGridDocument.ts` (`createDocumentFromTemplate`).

## Selection & Editing

Selection follows the same conventions as Excel and Google Sheets, with `useSelection` owning the reactive state, `useGridSelectionGestures` translating mouse and keyboard events into intent, and `useInlineCellEditor` handling the inline input field. The model tracks an active cell, a primary rectangular range, and a list of additional disjoint ranges so cells, rows, and columns can be mixed in one multi-selection.

- **Cells.** Click selects a single cell. Drag, or Shift+click, extends the primary range from the original anchor. Ctrl/Cmd+click on an unselected cell pushes the previous range onto the additional-ranges list and starts a new primary range. Ctrl/Cmd+click on an already-selected cell removes it from the selection — any range containing the cell is split into up to four sub-rectangles so the data structure stays rectangular. The last selected cell cannot be deselected, mirroring Excel.
- **Rows and columns.** Clicking a row or column header selects the whole row or column. Shift+click on another header extends from the anchor to cover every row or column in between. Ctrl/Cmd+click on a header adds (or, on an already-selected line, splits and removes) a disjoint row or column band. The same multi-selection list holds cell ranges, row bands, and column bands at once, so you can express "this column, that row, except their intersection" in one selection.
- **Inline editing.** Double-click, F2, or starting to type opens the inline editor — a `<textarea>` that auto-grows over the cell as the draft gets taller. Enter and Shift+Enter commit and move the active cell down or up; Tab and Shift+Tab commit and move right or left. **Alt+Enter (Option+Enter on macOS)** inserts a literal newline at the caret without committing, matching Excel's multi-line-cell gesture; the formula bar honours the same shortcut. Escape cancels. While typing a formula, clicking a different cell appends its A1 address to the editor instead of changing the selection.
- **Keyboard navigation.** Arrow keys, Tab, and Shift+Tab move the active cell outside the editor. They also clear additional ranges, matching how Excel collapses a multi-selection on a navigation keypress. Shift+arrow extends the primary range from the anchor.
- **Formatting and clipboard awareness.** Every selection-driven feature — Format cells, copy, cut, paste, the formula bar's address indicator — reads from the combined "all selected ranges" list, so a row + column + cells multi-selection is treated as one logical target.

## Excel Interop

TeamGrid imports and exports the .xlsx format through `exceljs`, and it copy/pastes ranges through the OS clipboard in an Excel-friendly way.

- **File import / export.** `File / Import XLSX` builds a TeamGrid workbook envelope from any `.xlsx` file, mapping values, supported formulas, number and date formats, fonts, fills, and alignment. `File / Export XLSX` produces an `.xlsx` blob from the current workbook. See `src/features/xlsx/lib/importWorkbook.ts` and `src/features/xlsx/lib/exportWorkbook.ts`.
- **Clipboard.** Copy and Cut emit three parallel encodings: a TeamGrid-native JSON payload embedded in the HTML clipboard for lossless TeamGrid-to-TeamGrid paste (preserves stable IDs, styles, and formula source); an Excel-compatible HTML body with `x:fmla`, `x:num`, and `x:str` attributes so Excel keeps relative formulas; and a plain TSV fallback for everything else. Paste decodes in the same priority order. See `src/features/clipboard/lib/`.

TeamGrid intentionally supports only a well-chosen subset of Excel's functionality — enough to demonstrate the collaboration model on a real-world structured document, not enough to replace Excel.

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
                  DeleteSpreadsheetDialog,RenameWorksheetDialog,
                  ViewSheetSettingsDialog}.vue
    view-sheets/                Virtual View Sheet materialization
      lib/viewSheetMaterialization.ts
    grid/                       spreadsheet surface and formula bar
      composables/{useSelection,useFormulaBarEditing,useFormulaAssistRouter,
                   useGridClipboard,useCellFormatDialog,useInlineCellEditor,
                   useGridSelectionGestures,useColumnRowResize,
                   useGridClipboardBridge}.ts
      components/{GridViewport,FormulaBar,FormulaAssistPanel,
                  WorksheetTabs,CellFormatDialog}.vue
      lib/{gridProjection,cellFormatting}.ts
    formulas/lib/               parser, evaluator, dependency tracking, registry
    charts/                     embedded charts (overlay, rendering, properties, Excel interop)
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
- `src/features/document/lib/viewOpen.ts` builds the three filter-mode dynamic view definitions (all / no templates / only templates) used by the File / Open and File / New from template... dialogs.
- `src/features/view-sheets/lib/viewSheetMaterialization.ts` walks a view navigator and projects its entries into the worksheet's stable-ID row, column, and cell maps so Virtual View Sheet rows persist alongside hand-edited cells.
- `src/features/grid/lib/gridProjection.ts` turns stable IDs into visible rows, columns, and addresses.
- `src/features/grid/lib/cellFormatting.ts` handles value coercion, date/number formatting, and style merging.
- `src/features/formulas/lib/` contains parsing, evaluation, dependency tracking, and function metadata.
- `src/features/clipboard/lib/` contains the serializer/deserializer for the TeamGrid + Excel + TSV clipboard payloads.
- `src/features/xlsx/lib/` contains the `.xlsx` import and export.
- `src/features/grid/components/GridViewport.vue` is the spreadsheet surface; it delegates selection, inline editing, column/row resizing, and native clipboard plumbing to the composables in `src/features/grid/composables/`.
- The dialog SFCs under `src/features/document/components/` and `src/features/grid/components/CellFormatDialog.vue` each take a controller object from their matching composable (`useErrorDialog`, `useOpenDialog`, `useDocumentPropertiesDialog`, `useWorksheetDialogs`, `useCellFormatDialog`) as a prop. The composable owns state and actions; the SFC renders the UI. `ViewSheetSettingsDialog.vue` follows the same pattern and is owned by `useWorksheetDialogs`.

## Current Scope

This sample is deliberately small. It does not yet implement virtualization for very large sheets, named ranges in the UI, cross-workbook formulas, conditional formatting, or every Excel formula and chart edge case. We are actively exploring more advanced editors — both open source and commercial — to edit Office formats in a data-sovereign way with full concurrency support, but most packages we evaluated are not yet powerful enough for our requirements (in particular: a clean separation between presentation and a stable-ID data model, and being able to drive edits through granular CRDT-friendly operations rather than whole-document replacements). The pieces above can all be layered onto TeamGrid's stable-ID schema and JSON-patch save path without changing the core collaboration model.
