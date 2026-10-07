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
| `src/xlsx/` | xlsx import (WASM sidecar) and export (vendored xlsx-gateway + JSZip). |
| `native/xlsx-wasm/` | wasm-bindgen wrapper around GenOffice's Rust xlsx reader. |

## Commands

```bash
pnpm install
pnpm build:wasm        # needs rustup toolchain 1.90+ with wasm32-unknown-unknown, and wasm-pack
pnpm dev               # http://localhost:4208
pnpm typecheck
pnpm sync-genoffice    # take a newer GenOffice checkout (../../../genoffice by default)
```

Open an xlsx with Cmd/Ctrl+O; Cmd/Ctrl+S downloads the edited workbook.

## Spike status

- Works: editor boot, formulas, xlsx open (styles, merges, charts, images) via
  WASM, edit, save as xlsx download, reopen.
- Not yet: MindooDB storage (schema v4), Haven SDK bridge, AutoSave, WASM in a
  Web Worker, TeamGrid features (templates, revisions, view sheets), WebMCP.
