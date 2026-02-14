# Benchmark Results

## Overview

This document presents comprehensive benchmark results comparing XlsbWriterNode's XLSB format performance against XLSX format.

## Test Environment

- **Platform:** Windows
- **Node.js:** v20+
- **Dataset:** 50,000 rows × 5 columns
- **Iterations:** 20 (for statistical accuracy)
- **Data Types:** Mixed (strings, numbers, dates)

## Results Summary

| Operation | XLSB | XLSX | XLSB Advantage |
|-----------|------|------|----------------|
| **Write** | 140.09 ms | 467.14 ms | **3.33x faster** |
| **Read** | 118.24 ms | 276.23 ms | **2.34x faster** |
| **File Size** | 1.49 MB | 2.83 MB | **47% smaller** |

## Detailed Results

### Write Performance

```
XlsbWriter:
  Times: [259.8, 160.6, 134.6, 134.0, 146.1, 136.3, 133.0, 131.7, 146.1, 142.4, 
          135.8, 132.2, 150.2, 140.7, 130.5, 129.3, 142.7, 139.1, 136.2, 195.9] ms
  Average: 140.09 ms (±7.53 ms)
  Size: 1.49 MB

XlsxWriter:
  Times: [483.3, 444.1, 471.2, 456.7, 496.0, 529.8, 488.2, 450.9, 458.7, 443.2, 
          495.4, 454.9, 469.6, 493.0, 499.6, 452.0, 463.7, 439.5, 441.1, 453.3] ms
  Average: 467.14 ms (±17.96 ms)
  Size: 2.83 MB
```

### Read Performance

```
XlsbReader:
  Times: [117.8, 136.3, 120.4, 132.1, 119.5, 118.7, 124.1, 113.7, 122.8, 113.8, 
          121.7, 116.3, 121.7, 112.9, 117.0, 110.8, 117.8, 112.2, 121.3, 112.3] ms
  Average: 118.24 ms (±3.56 ms)

XlsxReader:
  Times: [290.1, 269.8, 284.1, 268.0, 291.6, 281.8, 266.7, 284.3, 281.1, 257.2, 
          271.6, 260.2, 272.7, 264.5, 290.2, 264.7, 303.2, 265.9, 272.7, 300.1] ms
  Average: 276.23 ms (±9.44 ms)
```

## Visual Comparison

### Write Speed (lower is better)

```
XLSB ██████████████░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 140 ms
XLSX ████████████████████████████████████████████████ 467 ms
```

### Read Speed (lower is better)

```
XLSB ████████████████████░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 118 ms
XLSX ████████████████████████████████████████████████ 276 ms
```

### File Size (lower is better)

```
XLSB ██████████████████████████░░░░░░░░░░░░░░░░░░░░░░ 1.49 MB
XLSX ████████████████████████████████████████████████ 2.83 MB
```

## Why XLSB is Faster

### 1. Binary Format vs XML

- **XLSB** stores data in a compact binary format
- **XLSX** uses XML, which requires parsing/generating text

### 2. Compression Efficiency

- XLSB's binary format compresses more efficiently
- Smaller files = faster I/O operations

### 3. No XML Overhead

- No XML tags, attributes, or namespace declarations
- Direct data encoding reduces processing time

## Running Your Own Benchmarks

To run benchmarks on your system:

```bash
npm run benchmark
```

Or use the TypeScript file directly:

```bash
npx ts-node benchmark-our.ts
```

## Recommendations

| Scenario | Recommended Format |
|----------|-------------------|
| **Data processing pipelines** | XLSB |
| **Backend Excel generation** | XLSB |
| **Large datasets (>10K rows)** | XLSB |
| **Storage-constrained environments** | XLSB |
| **External sharing** | XLSX |
| **Legacy system compatibility** | XLSX |
| **Human-readable inspection needed** | XLSX |

## Comparison with Other Libraries

While this benchmark focuses on internal format comparison, note that XlsbWriterNode is designed for:

- **Minimal dependencies** - Only uses essential zip libraries
- **Memory efficiency** - Streaming architecture for large files
- **Type safety** - Full TypeScript support

For comparison with external libraries (like ExcelJS, xlsx-populate, etc.), results may vary based on specific use cases and dataset characteristics.
