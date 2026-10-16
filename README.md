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
pnpm build:wasm:node   # the same for Node (pkg-node/, committed too), which `pnpm test` uses for the xlsx round trips
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

## Revisions

File → Browse revisions lists the open workbook's revisions: the timelines
of the top document and its row blocks merged, the writes of one save (same
author, within 1.5 s) counted once. Picking one opens the workbook as every
document stood at that time, read-only and detached from the stored
document, with a banner leading back to the current state
(`src/haven/history.ts`). In Haven's time travel the whole app is read-only
and a banner names the date. The mock only keeps document history from
App SDK PR klehmann/mindoodb-app-sdk#20 on (`pnpm dev:local` with that
branch checked out).

## Agent tools

Inside Haven, TeamGrid offers tools to AI agents (WebMCP, prefix
`teamgrid_`, `src/agent/`). Reading and editing the open workbook goes through
GenOffice Sheets' own AI skill (`ai/tools.ts`), which the App publishes as
`window.teamGridSkillDeps` (`patches/sheets-agent-tools.patch`): the same
readers, the same workbook DSL and the same plan/apply path as its built-in
assistant, so an agent's batch is one undo step in the grid.

| Tool | What |
|---|---|
| `get_context` | open workbook, sheets with ids and extents, active sheet, selection |
| `workbooks_list`, `workbook_open`, `workbook_create` | stored workbooks and templates; create blank, from a template or from an imported .xlsx (`fileRef`) |
| `workbook_export` | .xlsx, or one sheet as .csv, as a `fileRef` for `haven_files_export` |
| `read_cells`, `read_range`, `read_formats`, `read_sheet_features`, `find_cells`, `aggregate_range`, `trace_precedents`, `trace_dependents`, `select_range` | GenOffice's readers |
| `apply_ops` | a batch in GenOffice's workbook DSL (cells, formulas, formats, layout, sheets, charts, shapes, pivots, tables, filters, conditional formats, validation, names); saved to MindooDB before it returns |
| `ops_guide` | GenOffice's op guides, per group |

After a write, opening or creating, the tools wait until the editor shows
the saved workbook and Univer finished recalculating, so the next read sees
formula results. In the test host: `__havenTestHost.agent.call('apply_ops', …)`.

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
- Stored with row/column ids: notes and hyperlinks (on their cells), the
  autoFilter with its criteria, print settings (orientation, paper, scale,
  fit, margin preset, gridlines/headings, header/footer, print area, print
  titles, page breaks) and tab colors (`src/model/sheet-features.ts`).
- Threaded comments (root text, author, replies, resolved flag) on their
  note; row/column outline levels and collapsed groups, summary position;
  sheet protection with allowed actions and password hash, allow-edit ranges
  without a password of their own (password ones are dropped, so their cells
  stay locked); veryHidden
  sheets (`src/model/build.ts`, `src/xlsx/tab-colors.ts`).
- Conditional formats and data validation: Univer's rule JSON (made from the
  file by GenOffice's own converters) with id-bound areas, keyed by content
  (`src/model/rules.ts`).
- Tables (style, header/totals row, banded rows/columns, first/last column,
  filter buttons; custom table styles stored workbook-wide with their formats,
  `src/xlsx/table-parts.ts`) and sparkline groups, with id-bound areas and
  sources (`src/model/objects.ts`).
- PivotTables: their xlsx parts (table, cache definition, records) as an
  attachment of the top document named by content hash, output area and
  source range by id; put back into every file built, with both references
  rewritten (`src/model/pivots.ts`, `src/xlsx/pivot-parts.ts`). Refresh in
  the editor updates the parts, the next save stores them.
- Charts: besides GenOffice's chart-add form, the file's chart part (with
  its style and colors parts) is kept as an attachment named by content
  hash and put back over the chart the export adds, its `<c:f>` references
  bound to ids (`src/xlsx/chart-parts.ts`). So chart types GenOffice cannot
  create (bubble, stock, surface) and all chart styling survive.
- Links to other workbooks: formulas keep `[n]Sheet!A1` verbatim, the
  externalLink parts (path and cached values) are stored as an attachment
  and written back in order (`src/xlsx/external-links.ts`).
- Slicers: caches, slicer parts, their extLst entries and drawing frames
  kept as an attachment (table ids and sheet tab ids as placeholders by
  name/stored id); the frame position is a visual with id-bound anchor
  (`src/xlsx/slicer-parts.ts`). Tested with table slicers; timelines are not
  kept.
- Macros (.xlsm): the VBA project with its signatures and the workbook/sheet
  code names are stored as an attachment; an export is again an .xlsm and
  downloads with that extension (`src/xlsx/macros.ts`). Macros never run in
  the browser.
- Not stored: allow-edit ranges with a password of their own, timelines.
- Not yet: WASM in a Web Worker, encryption,
  the embeddable "spreadsheet" component of TeamGrid 1.x, WebMCP.
