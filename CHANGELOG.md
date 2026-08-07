# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.1.0] - 2026-08-07

### Added

- `XlsxUpdater` and `XlsbUpdater` — replace the data of a worksheet inside an
  existing `.xlsx` / `.xlsb` file without rebuilding the workbook. Pivot tables,
  other sheets, styles and defined names are preserved; shared strings are
  appended (never reordered), column styles are inherited, and pivot cache
  ranges + `refreshOnLoad` are updated automatically.
- Shared helpers: `xmlUtils.ts` (XML escaping, column letters, shared strings
  parsing) and `biff12Utils.ts` (BIFF12 record helpers for XLSB).
- `examples/update-existing.ts` — end-to-end example for refreshing a data
  sheet in an existing XLSX / XLSB report.

### Fixed

- `XlsbUpdater` pivot source range was one row too long when `headers` were not
  provided, causing an empty trailing row (and a "(blank)" item in pivot
  tables). The range now ends exactly at the last data row.
- Both updaters now trim trailing rows that contain only empty cells (padding
  rows from SQL exports), so they don't appear as "(blank)" in pivot tables.
- `XlsbUpdater` no longer removes the worksheet's row-block terminator records
  (0x92/0x217/0x1dd/0x1dc) and keeps the sheet dimension record (0x98) in sync
  with the new row count — without these, Excel hid the last written row.

## [2.0.0] - 2026-07-25

### Breaking

- `XlsbReader.read()` is now `async` and returns `Promise<boolean>` (same as `XlsxReader`).
  Update loops from `while (reader.read())` to `while (await reader.read())`.
- Package entry points moved under `dist/cjs` and `dist/esm` with an `exports` map.
  Most consumers of `@justybase/spreadsheet-tasks` need no import path changes.

### Added

- `CellValue` / `PrimitiveCellValue` types for writer and reader cell APIs
- Shared writer helpers: `sanitizeSheetName`, column-width helpers, `StreamingSheetState`
- Dual package build (CJS + ESM) via `tsup`
- Browser sources under `src/browser/` with `npm run build:browser`
- `docs/PACKAGING.md` (ZIP library rationale + dual package notes)
- Dist surface smoke test (`npm run test:dist`)

### Changed

- `XlsbWriter` / `XlsxWriter` share sanitization, autofit, and streaming state helpers
- Tests import from `src/` (published surface still verified via `test:dist`)
- ZIP type definitions use `@types/*` only (hand-written ambient modules removed)

### Browser

- Source of truth is TypeScript in `src/browser/`; `browser/*.js` writer sources are generated bundles

## [1.3.2] - previous

See git history for 1.x changes.
