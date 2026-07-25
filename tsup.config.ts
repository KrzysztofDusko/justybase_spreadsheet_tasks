import { defineConfig } from 'tsup';

export default defineConfig([
    {
        entry: ['src/index.ts'],
        format: ['cjs'],
        outDir: 'dist/cjs',
        dts: true,
        sourcemap: true,
        clean: true,
        splitting: false,
        treeshake: false,
        target: 'node16',
        platform: 'node',
    },
    {
        entry: ['src/index.ts'],
        format: ['esm'],
        outDir: 'dist/esm',
        dts: false,
        sourcemap: true,
        clean: false,
        splitting: false,
        treeshake: false,
        target: 'node16',
        platform: 'node',
        outExtension() {
            return { js: '.js' };
        },
    },
]);
