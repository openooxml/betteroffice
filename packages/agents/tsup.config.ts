import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts', 'src/mcp.ts', 'src/render.ts', 'src/cli.ts'],
  format: ['esm'],
  dts: true,
  splitting: true,
  clean: true,
  external: [/^@betteroffice\//],
});
