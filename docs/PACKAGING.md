# Dependencies and packaging notes

## ZIP libraries

The Node package uses three ZIP libraries on purpose — they match different I/O models:

| Library | Used by | Why |
|---------|---------|-----|
| **archiver** | `XlsbWriter`, `XlsxWriter` | Streaming ZIP **write** into a file stream (low memory for large workbooks) |
| **adm-zip** | `XlsbReader` | Synchronous full-load ZIP **read** (XLSB worksheets are binary and parsed sync with BIFF) |
| **yauzl** | `XlsxReader` | Lazy async ZIP **read** (XLSX sheet XML can be large; entries loaded on demand) |

Unifying on a single ZIP library is intentionally out of scope until a measured migration (especially XLSB reader sync path) is justified.

Type definitions come from `@types/adm-zip`, `@types/archiver`, and `@types/yauzl` (no hand-written ambient modules).

## Dual package (CJS + ESM)

`npm run build` (tsup) emits:

- `dist/cjs/` — CommonJS (`require`)
- `dist/esm/` — ESM (`import`), with a local `package.json` `{ "type": "module" }`

Consumers should use the package root export:

```js
// CJS
const { XlsbWriter } = require('@justybase/spreadsheet-tasks');

// ESM
import { XlsbWriter } from '@justybase/spreadsheet-tasks';
```
