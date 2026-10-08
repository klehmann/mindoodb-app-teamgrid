# TeamGrid – new UI (spike)

GenOffice Sheets' editor (React + Univer OSS, Apache-2.0) running in the
browser, as the basis for TeamGrid's next version. Concept and decisions:
`GenOffice-Sheets-fuer-TeamGrid.md` in the project files.

## Layout

| Path | What |
|---|---|
| `vendor/genoffice/` | GenOffice sources, copied by `scripts/sync-genoffice.mjs` (commit in `VENDORED_FROM`). Never edit by hand. |
| `patches/` | Our changes to vendored files, applied by the sync script (`patch -p1` in `vendor/genoffice`). |
| `src/desktop-api-shim.ts` | Browser replacement for GenOffice's Electron bridge (`window.desktopApi`). |
| `src/xlsx/` | xlsx import (WASM sidecar) and export (vendored xlsx-gateway + JSZip); editor sessions. |
| `src/model/` | Stored workbook (schema v5): ids, formula references by id, row blocks, save diff → JSON patches. |
| `src/testing/` | Automerge-backed stand-in for MindooDB, for merge tests with two offline devices. |
| `src/haven/` | Haven App SDK connection and the workbook picker. |
| `native/xlsx-wasm/` | wasm-bindgen wrapper around GenOffice's Rust xlsx reader. |

## Commands

```bash
pnpm install
pnpm build:wasm        # needs rustup toolchain 1.90+ with wasm32-unknown-unknown, and wasm-pack
pnpm dev               # http://127.0.0.1:4208, mock Haven at /__haven-test/ (real Automerge
                       # documents; __havenTestHost.applyRemoteUpdate plays a second device)
pnpm typecheck
pnpm test             # model tests, incl. concurrent offline edits
pnpm build:wasm:node   # lets `pnpm test` also run the xlsx round trips (visuals)
pnpm sync-genoffice    # take a newer GenOffice checkout (../../../genoffice by default)
```

Inside Haven (or the mock at `/__haven-test/`) the editor opens stored
workbooks (form `teamgrid-next`), creates new ones and imports xlsx files as
documents; AutoSave writes JSON patches with `baseHeads`, and an idle editor
reloads when someone else changed the workbook. Standalone, Cmd/Ctrl+O opens
an xlsx and Cmd/Ctrl+S downloads it.

## How storage works

The editor only knows xlsx files, so a stored workbook is rendered to an xlsx
in memory (`model/export.ts`, GenOffice's gateway on a blank workbook) and
opened from there. On save, GenOffice patches that xlsx; it is read back in
full, rows and columns are matched to their ids by replaying the save's
inserts/deletes/moves (`model/replay.ts`), and the difference to the stored
document becomes one JSON patch (`model/sync.ts`). The editor then reopens
from the merged document, so concurrent changes show up right away.

A workbook is one top document (sheets, columns, merges, styles, the order
of row blocks) plus row-block documents of about 256 rows each, after the
pattern of mindoodb-word-journal. A save writes only the blocks it changed;
new rows join the block of the row before them, and appending past a full
block opens a new one with a derived id. Rows that only exist because the
sheet grew get ids derived from the row before them, so two people typing
into the next empty row offline end up in the same row
(`src/model/sync.test.ts` covers these cases in both merge directions).

A cell's format is stored per property (`s.bold`, `s.fillColor`, …), so
one person making a cell bold and another filling it yellow both keep their
change. Two sheets added offline under the same name, and merged areas that
overlap after a merge, are repaired the same way on every replica when the
workbook loads and written by the next save
(`src/model/merge-scenarios.test.ts`).

## View sheets

The Data tab's "MindooDB view" group adds a sheet filled from a virtual view
configured for the app in Haven (classic TeamGrid's virtual view sheet), and
refreshes or reconfigures the active one: a bold header row, category rows
on a light fill, documents below, dates as date cells. The sheet stores its
`viewBinding`; a refresh rewrites its cells, keeping row and column ids by
position (`src/model/view-sheet.ts`, `src/haven/view-sheets.ts`,
`patches/sheets-view-sheets.patch`). The test host configures a sample view,
"Kontakte nach Firma". Its mock hands out one navigator per view, so the top
level category is not applied there.

## Fill series

Univer's auto-fill only continues English and Chinese month and weekday names,
and its English lists overlap ("May" turns January…May into "Jun"). A rule of
our own (`src/fill-series.ts`, registered through
`patches/sheets-univer-hook.patch`) runs first: long and short names in the
eight TeamGrid languages (from `Intl`, plus Excel's German "Mrz"/"Mo" style),
the list that holds every source cell, the UI language first, and the source's
writing (JANUAR → FEBRUAR).

## Small screens

Below 1100px the tab row uses short labels and tabs that do not fit go
into "More"; ribbon groups that do not fit fold into dropdown buttons from
the right. On a phone (≤640px) a tab's commands open as a bottom sheet and
the ribbon starts collapsed (also in windows lower than 600px). The logic is
in `patches/sheets-responsive-ribbon.patch`, the styling in
`src/responsive.css`; it follows TeamSlides.

## Spike status

- Stored and round-tripped: values, formulas (references by row/column id),
  cell styles, row heights, column widths, hidden rows/columns, merges,
  frozen panes, gridlines, zoom, sheets (add, rename, hide, reorder, delete),
  charts (anchors and data ranges by id), pictures (bytes as attachments of
  the top document), basic shapes.
- Not stored yet (lost on reload): conditional formats, data validation,
  filters, notes, hyperlinks, tables, pivots, sparklines, page setup, defined
  names, tab colors; chart types GenOffice cannot create (bubble, stock,
  surface) and chart styling beyond its chart-add options.
- Not yet: WASM in a Web Worker, TeamGrid features (templates, revisions,
  view sheets, encryption), WebMCP, migration of v3 documents.
