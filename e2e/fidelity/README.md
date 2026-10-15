# Fidelity tests against Microsoft Excel

How close TeamGrid computes real Excel workbooks to Excel itself. The plan for the other
axes (rendering, round trip) is in `docs/fidelity-e2e-plan.md`.

## What runs

| Test | What it checks | Where |
| --- | --- | --- |
| `values.e2e.ts` | Each corpus workbook is imported into TeamGrid in the mock Haven (`workbook_create`). Every formula cell is read back with `read_cells`, and the value TeamGrid's formula engine computed is compared with the result Excel stored in the file. A workbook fails when a cell that matched at the baseline (`values-baseline.json`) no longer does. | CI and local |

The tests drive the app through its agent tools (`src/agent/`), so they exercise the
real import, store (MindooDB) and formula engine. They compare the engine's value, not the
displayed text: where the engine fails, TeamGrid displays Excel's cached result
(`formula-cached-fallback.ts`), which would hide the gap.

Comparison (`compare.ts`):
- Numbers may differ by 1e-9 relative (Excel keeps 15 significant digits).
- Text, booleans and error types must be equal.
- Cells calling volatile functions (`NOW`, `TODAY`, `RAND`, …) only must not be an error.

Workbooks too large for TeamGrid to load whole (streaming mode) keep Excel's stored
results without recalculating them; they are listed in the report and not compared. A
corpus entry with `skip` (a known failure such as TeamGrid stopping to respond) is
marked fixme.

State on 2026-10-15: 24,099 of 30,399 compared formula cells match (79.3 %), 90 of 215
workbooks and 179 of 414 functions without differences; 4 workbooks streamed, 2 skipped.

## Commands

```bash
pnpm test:fidelity                  # all workbooks, writes fidelity-report/index.html
pnpm test:fidelity -g @financial    # one category (tags in corpus.json)
pnpm test:fidelity -g ironcalc-edate
pnpm fidelity:accept                # last run's differences become the baseline
pnpm fidelity:fetch                 # download the corpus again (pinned commit)
```

`fidelity-report/index.html` lists the match rate per category and per Excel function,
the functions with differences first, and per workbook the cells that differ (Excel's
result, TeamGrid's, and whether the function is missing, errs or computes another value).
That list is the work list for formula gaps; fix one, run the tests, accept the baseline.

## Corpus

`corpus.json` lists the 221 workbooks of
[IronCalc's `xlsx/tests/calc_tests`](https://github.com/ironcalc/IronCalc/tree/62b3d81075236d3d987dfd742026d1d493a8fd20/xlsx/tests/calc_tests)
(MIT or Apache-2.0, license in `corpus/LICENSE-ironcalc`, hashes in
`corpus/SHA256SUMS`): one workbook per function group, saved by Excel 16, about 72,000
formula cells calling 428 functions. Excel for Mac 16.115 recalculating sample
workbooks reproduced every stored result, so the files serve as the oracle without
running Excel. Add a source or workbook by adding it to `corpus.json` and running
`pnpm fidelity:fetch`, `pnpm test:fidelity` and `pnpm fidelity:accept`.
