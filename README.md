# MindooDB Teamgrid

Teamgrid is a collaborative spreadsheet sample app for Haven and MindooDB. It demonstrates how to store one spreadsheet workbook in one MindooDB Automerge document while keeping rows, columns, worksheet tabs, cells, formulas, and formatting stable under collaboration.

The app intentionally favors clear architecture over feature completeness. It is meant to be read by platform developers who want to build document-style Haven apps.

## Development

```sh
pnpm install
pnpm --filter mindoodb-app-teamgrid dev
pnpm --filter mindoodb-app-teamgrid test
pnpm --filter mindoodb-app-teamgrid build
```

Use `pnpm --filter mindoodb-app-teamgrid dev:local` when you want Vite aliases to point at local MindooDB and App SDK source packages.

## Haven Capabilities

Teamgrid follows the same host bridge pattern as TeamEdit:

- `read` lists and opens spreadsheet documents.
- `create` enables File / New.
- `update` enables save and grid mutations.
- `delete` enables document deletion.
- `history` enables the revision picker and read-only historical snapshots.
- `attachments` is not used by the first spreadsheet sample.

Host time travel and document revisions are read-only. The app still allows sheet navigation, selection, copy-friendly viewing, and formula inspection in those modes.

## Document Model

Each MindooDB document stores one Teamgrid workbook:

- `subject` is the document title shown in file pickers.
- `kind` is `mindoodb.teamgrid`.
- `teamgrid` contains the versioned workbook schema.

Inside the workbook, worksheet tabs, rows, columns, and cells all use stable app-level IDs. The UI projects those IDs into Excel-style labels such as `A1`, but formulas persist references to row and column IDs. This prevents row or column insertions from silently retargeting formulas.

Rows, columns, and worksheets use tombstones instead of immediate hard deletion. This keeps historical references explainable and lets formulas render missing targets as errors rather than losing identity.

## Formulas

The formula subsystem is permissive-license sample code rather than a GPL/commercial engine integration. It currently supports:

- Arithmetic with `+`, `-`, `*`, and `/`.
- Cell references such as `A1`.
- Ranges such as `A1:B4`.
- `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `CONCAT`, and `TODAY`.
- A dependency graph for invalidation.
- Content assist with function signatures and inline help.
- Highlighting referenced cells and ranges while editing formulas.

The parser stores formula source text plus normalized stable-ID references. The function registry in `src/lib/formulas/functionRegistry.ts` is the shared source of truth for evaluation and formula authoring assistance.

## Formatting

Cells separate semantic values from presentation:

- Values can be empty, string, number, or date.
- Number and date values can carry simple display formats.
- Styles include text color, background color, font family, font size, bold, italic, underline, and alignment.
- Row and column defaults are kept separate from per-cell overrides so broad formatting does not rewrite every cell.

## Code Layout

- `src/composables/useTeamGridDocument.ts` owns Haven bridge setup, database/document lifecycle, capability gates, time travel, and revision snapshots.
- `src/lib/teamgridDocument.ts` defines the persisted schema, migrations, ID helpers, and default workbook factory.
- `src/lib/gridProjection.ts` turns stable IDs into visible rows, columns, and addresses.
- `src/lib/cellFormatting.ts` handles value coercion, date/number formatting, and style merging.
- `src/lib/formulas/` contains parsing, evaluation, dependency tracking, and function metadata.
- `src/components/GridViewport.vue`, `FormulaBar.vue`, and `WorksheetTabs.vue` make up the spreadsheet UI.

## Current Scope

This first sample is deliberately small. It does not yet implement import/export, virtualization for very large sheets, advanced Excel-compatible formats, cross-workbook formulas, or a full Excel formula language. Those can be layered onto the stable-ID schema without changing the core collaboration model.
