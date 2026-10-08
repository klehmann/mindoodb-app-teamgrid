# Mindoo TeamGrid

Collaborative spreadsheets for MindooDB Haven. The editor is GenOffice
Sheets' (React + Univer OSS, Apache-2.0) running in the browser; workbooks
are stored as Automerge-friendly MindooDB documents. TeamGrid 1.x (the Vue
app with its own grid) lives in the git history before version 2.0.0.

## Layout

| Path | What |
|---|---|
| `vendor/genoffice/` | GenOffice sources, copied by `scripts/sync-genoffice.mjs` (commit in `VENDORED_FROM`). Never edit by hand. |
| `patches/` | Our changes to vendored files, applied by the sync script (`patch -p1` in `vendor/genoffice`). |
| `src/desktop-api-shim.ts` | Browser replacement for GenOffice's Electron bridge (`window.desktopApi`). |
| `src/xlsx/` | xlsx import (WASM sidecar) and export (vendored xlsx-gateway + JSZip); editor sessions. |
| `src/model/` | Stored workbook (schema v5): ids, formula references by id, row blocks, save diff → JSON patches. |
| `src/testing/` | Automerge-backed stand-in for MindooDB, for merge tests with two offline devices. |
| `src/haven/` | Haven App SDK connection, welcome screen, file menu, dialogs, view sheets. |
| `src/boot.ts` | Landing page when opened outside Haven (`?standalone` opens the editor anyway). |
| `native/xlsx-wasm/` | wasm-bindgen wrapper around GenOffice's Rust xlsx reader; `pkg/` is committed, so building the app needs no Rust toolchain. |
| `public/` | App definition (`haven-app.json`), store listing, Cloudflare headers and 404 page; `sw.js` removes TeamGrid 1.x's service worker. |

## Commands

```bash
pnpm install
pnpm dev               # http://127.0.0.1:4207, mock Haven at /__haven-test/ (real Automerge
                       # documents; __havenTestHost.applyRemoteUpdate plays a second device)
pnpm dev:local         # the same with mindoodb, the App SDK and the view language from their
                       # sibling checkouts' sources (../mindoodb, ../mindoodb-app-sdk, …)
pnpm build             # typecheck + production build into dist/ (incl. haven-bundle.zip)
pnpm deploy            # build and deploy to Cloudflare (wrangler)
pnpm build:wasm        # rebuild native/xlsx-wasm/pkg after changing the Rust wrapper; needs rustup
                       # toolchain 1.90+ with wasm32-unknown-unknown, and wasm-pack
pnpm typecheck
pnpm test             # model tests, incl. concurrent offline edits
pnpm build:wasm:node   # lets `pnpm test` also run the xlsx round trips (visuals)
pnpm sync-genoffice    # take a newer GenOffice checkout (../../genoffice by default)
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

## TeamGrid 1.x workbooks

Workbooks of TeamGrid 1.x (form `teamgrid`) are listed next to the current
ones. Opening one creates a copy in the current format (form
`teamgrid-next`, field `copiedFrom` = the original's id) and opens that; the
original is never written. From then on the list shows the copy instead of
the original. The conversion keeps all ids, so formulas, chart ranges and
view sheets stay bound (`src/model/legacy.ts`, `src/haven/legacy-copy.ts`).
The test host seeds a 1.x sample workbook ("Umsatz 2025").

## Small screens

Below 1100px the tab row uses short labels and tabs that do not fit go
into "More"; ribbon groups that do not fit fold into dropdown buttons from
the right. On a phone (≤640px) a tab's commands open as a bottom sheet and
the ribbon starts collapsed (also in windows lower than 600px). The logic is
in `patches/sheets-responsive-ribbon.patch`, the styling in
`src/responsive.css`; it follows TeamSlides.

## Status

- Stored and round-tripped: values, formulas (references by row/column id),
  cell styles, row heights, column widths, hidden rows/columns, merges,
  frozen panes, gridlines, zoom, sheets (add, rename, hide, reorder, delete),
  charts (anchors and data ranges by id), pictures (bytes as attachments of
  the top document), basic shapes.
- Defined names are stored by key (scope + lower-case name) with id-bound
  formulas, and written back on every open; formula cells are exported with
  their last result as cached value.
- Not stored yet (lost on reload): conditional formats, data validation,
  filters, notes, hyperlinks, tables, pivots, sparklines, page setup, tab
  colors; chart types GenOffice cannot create (bubble, stock,
  surface) and chart styling beyond its chart-add options.
- Not yet: WASM in a Web Worker, revisions (Haven time travel), encryption,
  the embeddable "spreadsheet" component of TeamGrid 1.x, WebMCP.
