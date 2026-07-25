# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
